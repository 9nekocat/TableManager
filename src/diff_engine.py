# -*- coding: utf-8 -*-
"""
Diff engine for DataTable Manager.
Reuses merge_engine's workbook reader for correct formula handling.
Compare current working copy of a table with a specific SVN revision.
"""

import tempfile
import os
from collections import OrderedDict
from pathlib import Path
from difflib import SequenceMatcher

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from manager_core import RAW_DIR, TEMP_DIR, DATA_START_ROW
from merge_engine import (
    _external_link_index_map,
    _normalize_external_formula_refs,
    _read_all_rows_with_formulas,
    MergePreview,
    align_datasets_by_headers,
    _header_signatures,
)
from svn_client import SvnClient, SvnError
from excel_recalc import recalc_to_temp, remove_with_retry, schedule_deletion

# Styling
_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
_HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
_ADD_FILL = PatternFill(start_color="DDFFDD", end_color="DDFFDD", fill_type="solid")
_DEL_FILL = PatternFill(start_color="FFDDDD", end_color="FFDDDD", fill_type="solid")
_MOD_OLD_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
_MOD_NEW_FILL = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
_LABEL_FONT = Font(bold=True, color="333333")
_CHANGED_FONT = Font(bold=True, color="CC0000")
_CHANGED_NEW_FONT = Font(bold=True, color="006600")
_COLUMN_FILL = PatternFill(start_color="EADCF8", end_color="EADCF8", fill_type="solid")
_COLUMN_FONT = Font(bold=True, color="5A2D82")

_COLUMN_CHANGE_TYPES = {
    "column_added",
    "column_deleted",
    "column_modified",
    "column_moved",
}


def _is_column_change(change):
    return change.get("type") in _COLUMN_CHANGE_TYPES


def _column_name(meta):
    if not meta:
        return ""
    return (
        str(meta.get("row2") or "").strip()
        or str(meta.get("row1") or "").strip()
        or f"col{meta.get('source_col', '')}"
    )


def _column_header_tuple(meta):
    if not meta:
        return ("", "", "")
    return (
        str(meta.get("row1") or "").strip(),
        str(meta.get("row2") or "").strip(),
        str(meta.get("row3") or "").strip(),
    )


def _column_change(
    change_type,
    *,
    old_meta=None,
    new_meta=None,
    aligned_sig_to_col=None,
):
    aligned_sig_to_col = aligned_sig_to_col or {}
    cols = []
    for meta in (new_meta, old_meta):
        if meta and meta.get("sig") in aligned_sig_to_col:
            cols.append(aligned_sig_to_col[meta["sig"]])
    cols = list(OrderedDict.fromkeys(cols))
    display_name = _column_name(new_meta) or _column_name(old_meta)
    old_name = _column_name(old_meta)
    new_name = _column_name(new_meta)
    return {
        "type": change_type,
        "rowname": display_name,
        "col": cols[0] if cols else None,
        "cols": cols,
        "old_col": old_meta.get("source_col") if old_meta else None,
        "new_col": new_meta.get("source_col") if new_meta else None,
        "old_name": old_name,
        "new_name": new_name,
        "old_header": _column_header_tuple(old_meta),
        "new_header": _column_header_tuple(new_meta),
    }


def _compare_column_headers(old_headers, new_headers, aligned_headers):
    """Return column-level changes keyed by sheet.

    Row comparison receives these changes separately so column additions and
    deletions do not make every data row look modified.
    """
    result = {}
    all_sheets = set(old_headers.keys()) | set(new_headers.keys()) | set(aligned_headers.keys())

    for sn in all_sheets:
        old_metas = list(_header_signatures(old_headers.get(sn, {})).values())
        new_metas = list(_header_signatures(new_headers.get(sn, {})).values())
        aligned_sig_to_col = {
            meta["sig"]: col
            for col, meta in _header_signatures(aligned_headers.get(sn, {})).items()
        }

        old_by_sig = {meta["sig"]: meta for meta in old_metas}
        new_by_sig = {meta["sig"]: meta for meta in new_metas}
        old_sigs = [meta["sig"] for meta in old_metas]
        new_sigs = [meta["sig"] for meta in new_metas]
        old_sig_set = set(old_sigs)
        new_sig_set = set(new_sigs)
        deleted_sigs = old_sig_set - new_sig_set
        added_sigs = new_sig_set - old_sig_set

        changes = []
        paired_deleted = set()
        paired_added = set()

        matcher = SequenceMatcher(a=old_sigs, b=new_sigs, autojunk=False)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != "replace":
                continue
            old_candidates = [meta for meta in old_metas[i1:i2] if meta["sig"] in deleted_sigs]
            new_candidates = [meta for meta in new_metas[j1:j2] if meta["sig"] in added_sigs]
            for old_meta, new_meta in zip(old_candidates, new_candidates):
                changes.append(_column_change(
                    "column_modified",
                    old_meta=old_meta,
                    new_meta=new_meta,
                    aligned_sig_to_col=aligned_sig_to_col,
                ))
                paired_deleted.add(old_meta["sig"])
                paired_added.add(new_meta["sig"])

        for sig in old_sigs:
            if sig in deleted_sigs and sig not in paired_deleted:
                changes.append(_column_change(
                    "column_deleted",
                    old_meta=old_by_sig[sig],
                    aligned_sig_to_col=aligned_sig_to_col,
                ))

        for sig in new_sigs:
            if sig in added_sigs and sig not in paired_added:
                changes.append(_column_change(
                    "column_added",
                    new_meta=new_by_sig[sig],
                    aligned_sig_to_col=aligned_sig_to_col,
                ))

        moved_sigs = set()
        common_sigs = old_sig_set & new_sig_set
        old_common = [sig for sig in old_sigs if sig in common_sigs]
        new_common = [sig for sig in new_sigs if sig in common_sigs]
        common_matcher = SequenceMatcher(a=old_common, b=new_common, autojunk=False)
        for tag, i1, i2, j1, j2 in common_matcher.get_opcodes():
            if tag == "equal":
                continue
            moved_sigs.update(old_common[i1:i2])
            moved_sigs.update(new_common[j1:j2])

        for sig in common_sigs:
            old_meta = old_by_sig[sig]
            new_meta = new_by_sig[sig]
            if _column_header_tuple(old_meta) != _column_header_tuple(new_meta):
                changes.append(_column_change(
                    "column_modified",
                    old_meta=old_meta,
                    new_meta=new_meta,
                    aligned_sig_to_col=aligned_sig_to_col,
                ))
            if sig in moved_sigs:
                changes.append(_column_change(
                    "column_moved",
                    old_meta=old_meta,
                    new_meta=new_meta,
                    aligned_sig_to_col=aligned_sig_to_col,
                ))

        changes.sort(key=lambda c: (
            min(c.get("cols") or [9999]),
            c.get("type") or "",
            c.get("new_col") or c.get("old_col") or 9999,
        ))
        if changes:
            result[sn] = changes

    return result


def _extract_value(cell_tuple):
    """Extract computed value from (value, formula) tuple.
    Falls back to formula text when cached value is None
    (e.g., after openpyxl save/load cycle loses Excel cached values).
    This prevents false-positive diffs on unchanged formula cells.

    The ' → ' split only applies to merge-preview display format
    ('=FORMULA → VALUE, optionally without the apostrophe once stored) —
    a plain config string like "A→B" is NOT a display format and must
    be kept as-is."""
    if cell_tuple is None:
        return ""
    val, fml = cell_tuple
    if val is not None:
        s = str(val)
        arrow = chr(0x2192)  # →
        if arrow in s and s.startswith(("'=", "=")):
            s = s.split(arrow, 1)[1].strip()
        return s
    if fml is not None:
        return fml
    return ""


def _cell_display(cell_tuple):
    """Convert (value, formula) tuple to display string — value only."""
    if cell_tuple is None:
        return ""
    val, _ = cell_tuple
    if val is None:
        return ""
    return _extract_value(cell_tuple)


def _cell_compare_str(cell_tuple):
    """Get a stable comparison string from (value, formula) tuple.
    Compare by computed value only — never compare formulas."""
    return _extract_value(cell_tuple)


def _normalize_dataset_formula_refs(data, workbook_path):
    """Normalize external-link formula slots in a read dataset in-place."""
    index_map = _external_link_index_map(workbook_path)
    if not index_map:
        return data
    for rows in data.values():
        for row_data in rows.values():
            for col_idx, cell in list(row_data.items()):
                if not cell:
                    continue
                val, fml = cell
                if isinstance(fml, str):
                    row_data[col_idx] = (val, _normalize_external_formula_refs(fml, index_map))
    return data


def _get_old_version(file_path, revision):
    """Export file at revision to temp and read directly.
    Old SVN versions have Excel-cached formula values — no recalc needed."""
    svn = SvnClient()
    rev_str = str(revision).lstrip("r")
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()

    try:
        svn.cat_to_file(file_path, rev_str, tmp.name)
        data, headers = _read_all_rows_with_formulas(tmp.name)
        _normalize_dataset_formula_refs(data, tmp.name)
    except SvnError as e:
        raise RuntimeError(f"无法获取 r{rev_str} 版本: {e}")
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass

    return data, headers


def _read_current_version_for_diff(full_path):
    """Read the current workbook after refreshing formula caches on a copy.

    openpyxl preserves formulas but cannot write Excel's cached formula values.
    A file that was just resolved by the merge engine may therefore compare as
    empty/stale until Excel opens and saves it. Recalculate a same-directory
    copy before reading, then remove the copy.
    """
    recalc_path = None
    read_path = str(full_path)
    try:
        recalc_path = recalc_to_temp(str(full_path), TEMP_DIR)
        read_path = recalc_path
    except Exception:
        recalc_path = None

    try:
        data, headers = _read_all_rows_with_formulas(read_path)
        _normalize_dataset_formula_refs(data, read_path)
        return data, headers
    finally:
        if recalc_path and recalc_path != str(full_path):
            if not remove_with_retry(recalc_path):
                schedule_deletion(recalc_path)


def _compare_data(old_data, new_data, column_changes=None):
    """
    Compare two datasets from _read_all_rows_with_formulas.
    Format: {sheet: {rowname: {col: (value, formula)}}}
    Returns: {sheet: {"status", "changes": [...]}}
    """
    column_changes = column_changes or {}
    all_sheets = sorted(set(list(old_data.keys()) + list(new_data.keys()) + list(column_changes.keys())))
    result = {}

    for sn in all_sheets:
        old_rows = old_data.get(sn, OrderedDict())
        new_rows = new_data.get(sn, OrderedDict())
        sheet_column_changes = list(column_changes.get(sn, []))
        column_level_cols = set()
        for change in sheet_column_changes:
            if change.get("type") in ("column_added", "column_deleted", "column_modified"):
                column_level_cols.update(col for col in change.get("cols", []) if col)

        if not old_rows and new_rows:
            result[sn] = {"status": "added", "changes": sheet_column_changes}
            continue
        if old_rows and not new_rows:
            result[sn] = {"status": "deleted", "changes": sheet_column_changes}
            continue

        changes = list(sheet_column_changes)
        all_keys = list(OrderedDict.fromkeys(
            list(old_rows.keys()) + list(new_rows.keys())
        ))

        for key in all_keys:
            old_row = old_rows.get(key)
            new_row = new_rows.get(key)

            if old_row is None and new_row is not None:
                changes.append({"rowname": key, "type": "added"})
            elif old_row is not None and new_row is None:
                changes.append({"rowname": key, "type": "deleted"})
            elif old_row is not None and new_row is not None:
                all_cols = sorted(set(list(old_row.keys()) + list(new_row.keys())))
                diffs = []
                for col in all_cols:
                    if col in column_level_cols:
                        continue
                    old_cell = old_row.get(col)
                    new_cell = new_row.get(col)
                    old_c = _cell_compare_str(old_cell)
                    new_c = _cell_compare_str(new_cell)
                    if old_c != new_c:
                        # Suppress false positives from lost cached formula values:
                        # if both sides share the same formula and one side is
                        # missing its cached value (saved by openpyxl), the diff
                        # is just a cache artifact. When BOTH sides have real
                        # cached values, a value difference is genuine (e.g.
                        # external dependency changed) and must be reported.
                        old_fml = old_cell[1] if old_cell and len(old_cell) > 1 else None
                        new_fml = new_cell[1] if new_cell and len(new_cell) > 1 else None
                        old_val = old_cell[0] if old_cell else None
                        new_val = new_cell[0] if new_cell else None
                        if old_fml is not None and old_fml == new_fml:
                            if old_val is None or new_val is None:
                                continue
                        # Both sides recalculated to EMPTY. After the pre-diff
                        # recalc a None cache means the formula result is empty
                        # (Excel caches "" as t="str" <v/>, which openpyxl reads
                        # as None), and a cell without formula is simply empty.
                        # Empty vs empty is not a change — even when one side
                        # gained a formula, or both formulas differ only
                        # because rows above shifted (e.g. =B449 vs =B450).
                        if old_val is None and new_val is None:
                            continue
                        diffs.append({
                            "col": col,
                            "old": str(_cell_display(old_cell)),
                            "new": str(_cell_display(new_cell)),
                        })
                if diffs:
                    changes.append({"rowname": key, "type": "modified", "cell_diffs": diffs})

        result[sn] = {
            "status": "modified" if changes else "same",
            "changes": changes,
        }

    return result


def _visible_diff_columns(sheet_diff, old_rows, new_rows, max_col):
    """Return data-column indexes that should be visible in diff Excel.

    The workbook keeps all columns so header alignment remains debuggable, but
    unchanged fields are hidden to keep wide tables such as Card readable.
    """
    visible = {1, 2}
    for change in sheet_diff.get("changes", []):
        change_type = change.get("type")
        if _is_column_change(change):
            visible.update(
                col for col in change.get("cols", [])
                if 1 <= col <= max_col
            )
        elif change_type == "modified":
            visible.update(d["col"] for d in change.get("cell_diffs", []))
        elif change_type == "added":
            row = new_rows.get(change.get("rowname"), {})
            visible.update(
                col for col, cell in row.items()
                if 1 <= col <= max_col and _cell_display(cell) != ""
            )
        elif change_type == "deleted":
            row = old_rows.get(change.get("rowname"), {})
            visible.update(
                col for col, cell in row.items()
                if 1 <= col <= max_col and _cell_display(cell) != ""
            )
    return {col for col in visible if 1 <= col <= max_col}


def generate_diff(file_path, revision):
    """
    Compare a table file's current version with a given SVN revision.
    Returns structured diff data for the web UI.
    """
    full_path = RAW_DIR / file_path
    if not full_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    rev_str = str(revision).lstrip("r")
    old_data, old_headers = _get_old_version(file_path, revision)

    # Read current version with formula caches refreshed on a temporary copy.
    new_data, current_headers = _read_current_version_for_diff(full_path)
    new_data, old_data, new_headers = align_datasets_by_headers(
        new_data, current_headers, old_data, old_headers
    )
    column_changes = _compare_column_headers(old_headers, current_headers, new_headers)

    sheets = _compare_data(old_data, new_data, column_changes)

    # Attach headers to each sheet for UI
    for sn, sd in sheets.items():
        hdr = new_headers.get(sn, {})
        sd["headers"] = [
            list(hdr.get("row1", ())),
            list(hdr.get("row2", ())),
            list(hdr.get("row3", ())),
        ]

    # Summary stats
    total_added = 0
    total_deleted = 0
    total_modified = 0
    column_added = 0
    column_deleted = 0
    column_modified = 0
    column_moved = 0
    for sd in sheets.values():
        for c in sd.get("changes", []):
            if c["type"] == "added":
                total_added += 1
            elif c["type"] == "deleted":
                total_deleted += 1
            elif c["type"] == "modified":
                total_modified += 1
            elif c["type"] == "column_added":
                column_added += 1
                total_modified += 1
            elif c["type"] == "column_deleted":
                column_deleted += 1
                total_modified += 1
            elif c["type"] == "column_modified":
                column_modified += 1
                total_modified += 1
            elif c["type"] == "column_moved":
                column_moved += 1
                total_modified += 1

    return {
        "file": file_path,
        "revision": f"r{rev_str}",
        "summary": {
            "added": total_added,
            "deleted": total_deleted,
            "modified": total_modified,
            "column_added": column_added,
            "column_deleted": column_deleted,
            "column_modified": column_modified,
            "column_moved": column_moved,
        },
        "sheets": sheets,
    }


def _build_diff_excel(sheets_diff, old_data, new_data, new_headers, output_path, old_label, new_label):
    """Write diff Excel workbook. Shared by export_diff_to_excel and diff_two_files."""
    out_wb = openpyxl.Workbook()
    first_sheet = True

    for sn in sorted(sheets_diff.keys()):
        sd = sheets_diff[sn]
        if sd["status"] == "same":
            continue

        if first_sheet:
            ws = out_wb.active
            ws.title = sn
            first_sheet = False
        else:
            ws = out_wb.create_sheet(title=sn)

        hdr = new_headers.get(sn, {})
        max_col = hdr.get("max_col", 1)

        old_rows = old_data.get(sn, OrderedDict())
        new_rows = new_data.get(sn, OrderedDict())
        visible_cols = _visible_diff_columns(sd, old_rows, new_rows, max_col)

        _write_header_row(ws, 1, "[标记]", hdr.get("row1", ()), max_col)
        _write_header_row(ws, 2, "(EN)", hdr.get("row2", ()), max_col)
        _write_header_row(ws, 3, "(type)", hdr.get("row3", ()), max_col)

        current_row = DATA_START_ROW

        changes = sd.get("changes", [])
        for idx, change in enumerate(changes):
            rowname = change["rowname"]

            if _is_column_change(change):
                _write_column_change_row(ws, current_row, change, max_col)
                current_row += 1
                next_change = changes[idx + 1] if idx + 1 < len(changes) else None
                if next_change and not _is_column_change(next_change):
                    current_row += 1

            elif change["type"] == "added":
                row_data = new_rows.get(rowname, {})
                _write_diff_row(ws, current_row, "[新增]", row_data, max_col, _ADD_FILL, None)
                current_row += 2

            elif change["type"] == "deleted":
                row_data = old_rows.get(rowname, {})
                _write_diff_row(ws, current_row, "[删除]", row_data, max_col, _DEL_FILL, None)
                current_row += 2

            elif change["type"] == "modified":
                changed_cols = set(d["col"] for d in change.get("cell_diffs", []))

                old_row_data = old_rows.get(rowname, {})
                _write_diff_row(ws, current_row, old_label, old_row_data, max_col, _MOD_OLD_FILL, changed_cols, _CHANGED_FONT)
                current_row += 1

                new_row_data = new_rows.get(rowname, {})
                _write_diff_row(ws, current_row, new_label, new_row_data, max_col, _MOD_NEW_FILL, changed_cols, _CHANGED_NEW_FONT)
                current_row += 1

                current_row += 1

        ws.column_dimensions["A"].width = 14
        for ci in range(max_col):
            data_col = ci + 1
            excel_col = get_column_letter(ci + 2)
            ws.column_dimensions[excel_col].width = 18
            if data_col not in visible_cols:
                ws.column_dimensions[excel_col].hidden = True
        ws.freeze_panes = f"A{DATA_START_ROW}"

    if first_sheet:
        ws = out_wb.active
        ws.title = "NoDiff"
        ws.cell(row=1, column=1, value="无差异")

    out_wb.save(str(output_path))
    return str(output_path)


def export_diff_to_excel(file_path, revision):
    """Generate diff Excel comparing current file vs SVN revision."""
    import uuid
    full_path = RAW_DIR / file_path
    if not full_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    rev_str = str(revision).lstrip("r")
    old_data, old_headers = _get_old_version(file_path, revision)

    # Read current version with formula caches refreshed on a temporary copy.
    new_data, current_headers = _read_current_version_for_diff(full_path)
    new_data, old_data, new_headers = align_datasets_by_headers(
        new_data, current_headers, old_data, old_headers
    )
    column_changes = _compare_column_headers(old_headers, current_headers, new_headers)

    sheets_diff = _compare_data(old_data, new_data, column_changes)

    stem = Path(file_path).stem
    rel_parent = Path(file_path).parent
    output_dir = TEMP_DIR / rel_parent if str(rel_parent) != "." else TEMP_DIR
    os.makedirs(output_dir, exist_ok=True)
    uid = uuid.uuid4().hex[:6]
    output_path = output_dir / f"{stem}_diff_r{rev_str}_{uid}.xlsx"

    return _build_diff_excel(sheets_diff, old_data, new_data, new_headers,
                             output_path, f"[旧版 r{rev_str}]", "[当前]")


def diff_two_files(old_path, new_path, output_dir, display_path=None):
    """Compare two xlsx files directly, export diff to Excel.
    Designed for TortoiseSVN external diff tool integration.
    Returns output path and opens the file in Excel."""
    import shutil as _shutil
    import uuid
    from excel_recalc import recalc_batch, cleanup_stale_recalc_copies, remove_with_retry

    new_p = Path(new_path)
    display_p = Path(display_path) if display_path else new_p
    copy_dir = display_p.parent if display_p.parent.exists() else new_p.parent
    stem = display_p.stem or new_p.stem or "workbook"
    # Sweep leftover recalc copies from previous crashed/failed runs
    cleanup_stale_recalc_copies(copy_dir)
    # Unique suffix prevents collision across consecutive runs
    uid = uuid.uuid4().hex[:6]
    old_tmp = os.path.join(str(copy_dir), f".{stem}_old_recalc_{uid}.xlsx")
    new_tmp = os.path.join(str(copy_dir), f".{stem}_new_recalc_{uid}.xlsx")

    try:
        _shutil.copy2(old_path, old_tmp)
        _shutil.copy2(new_path, new_tmp)

        # Batch recalc both copies in one Excel session
        recalc_batch([old_tmp, new_tmp])

        old_data, old_headers = _read_all_rows_with_formulas(old_tmp)
        new_data, current_headers = _read_all_rows_with_formulas(new_tmp)
        _normalize_dataset_formula_refs(old_data, old_tmp)
        _normalize_dataset_formula_refs(new_data, new_tmp)
    finally:
        for tmp in (old_tmp, new_tmp):
            # Excel may keep the just-recalc'd copies locked briefly after
            # Quit(); retry patiently, then hand off to a background reaper
            # so no temp copy is left in the working-copy directory.
            if not remove_with_retry(tmp, attempts=10, delay=0.5):
                schedule_deletion(tmp)

    new_data, old_data, new_headers = align_datasets_by_headers(
        new_data, current_headers, old_data, old_headers
    )
    column_changes = _compare_column_headers(old_headers, current_headers, new_headers)

    sheets_diff = _compare_data(old_data, new_data, column_changes)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(str(output_dir), f"{stem}_diff_{uid}.xlsx")

    result = _build_diff_excel(sheets_diff, old_data, new_data, new_headers,
                               output_path, "[旧版]", "[当前]")

    os.startfile(str(result))
    return result


def _write_header_row(ws, row, label, header_tuple, max_col):
    """Write a styled header row."""
    c = ws.cell(row=row, column=1, value=label)
    c.fill = _HEADER_FILL
    c.font = _HEADER_FONT
    for ci in range(max_col):
        val = header_tuple[ci] if ci < len(header_tuple) else ""
        c = ws.cell(row=row, column=ci + 2, value=val)
        c.fill = _HEADER_FILL
        c.font = _HEADER_FONT


def _column_change_label(change_type):
    return {
        "column_added": "[列新增]",
        "column_deleted": "[列删除]",
        "column_modified": "[列修改]",
        "column_moved": "[列移动]",
    }.get(change_type, "[列变更]")


def _column_change_message(change):
    change_type = change.get("type")
    old_name = change.get("old_name") or ""
    new_name = change.get("new_name") or ""
    old_col = change.get("old_col")
    new_col = change.get("new_col")
    if change_type == "column_added":
        return f"新增列: {new_name} (new col {new_col})"
    if change_type == "column_deleted":
        return f"删除列: {old_name} (old col {old_col})"
    if change_type == "column_modified":
        return f"列修改: {old_name} -> {new_name} (old col {old_col}, new col {new_col})"
    if change_type == "column_moved":
        return f"列移动: {new_name or old_name} (old col {old_col}, new col {new_col})"
    return "列变更"


def _write_column_change_row(ws, row, change, max_col):
    label_cell = ws.cell(row=row, column=1, value=_column_change_label(change.get("type")))
    label_cell.font = _COLUMN_FONT
    label_cell.fill = _COLUMN_FILL

    cols = [col for col in change.get("cols", []) if 1 <= col <= max_col]
    target_col = (cols[0] if cols else 1) + 1
    for excel_col in range(2, max_col + 2):
        cell = ws.cell(row=row, column=excel_col, value="")
        cell.fill = _COLUMN_FILL

    message_cell = ws.cell(row=row, column=target_col, value=_column_change_message(change))
    message_cell.font = _COLUMN_FONT
    message_cell.fill = _COLUMN_FILL
    for col in cols[1:]:
        extra_cell = ws.cell(row=row, column=col + 1, value=_column_change_message(change))
        extra_cell.font = _COLUMN_FONT
        extra_cell.fill = _COLUMN_FILL


def _write_diff_row(ws, row, label, row_data, max_col, fill, highlight_cols, highlight_font=None):
    """
    Write a data row with label.
    row_data: {col_idx: (value, formula)} from merge engine reader.
    """
    label_cell = ws.cell(row=row, column=1, value=label)
    label_cell.font = _LABEL_FONT
    label_cell.fill = fill

    for col_idx in range(1, max_col + 1):
        cell_tuple = row_data.get(col_idx)
        display = _cell_display(cell_tuple)
        c = ws.cell(row=row, column=col_idx + 1, value=display)
        c.fill = fill
        if highlight_cols and col_idx in highlight_cols and highlight_font:
            c.font = highlight_font
