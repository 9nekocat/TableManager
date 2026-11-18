# -*- coding: utf-8 -*-
"""
Merge engine for DataTable Manager.
- 3-way conflict detection (base / mine / theirs)
- Merge preview Excel generation with [我的]/[仓库]/[最终] layout
- Resolve: read decisions from preview, apply with formula adjustment
"""

import os
import copy
import sys
import zipfile
import tempfile
import shutil
import re
from urllib.parse import unquote
from collections import OrderedDict
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

from manager_core import RAW_DIR, DATA_START_ROW
from formula_utils import (
    format_display,
    adjust_formula,
    build_row_map,
    normalize_formula_for_compare,
)

# ---------------------------------------------------------------------------
# Cell status constants
# ---------------------------------------------------------------------------
SAME = "same"
MINE_MOD = "mine_modified"
THEIRS_MOD = "theirs_modified"
CONFLICT = "conflict"
MINE_ADDED = "mine_added"
THEIRS_ADDED = "theirs_added"
BOTH_ADDED = "both_added"
MINE_DELETED = "mine_deleted"
THEIRS_DELETED = "theirs_deleted"

DECISION_MODIFY = "[修改]"
DECISION_DELETE = "[删除]"
DECISION_CONFLICT_MINE = "mine"
DECISION_CONFLICT_THEIRS = "theirs"
DECISION_CONFLICT_EMPTY = "空"


def _normalize_conflict_default(value):
    """Normalize conflict priority choice for preview prefill."""
    text = str(value or "").strip().lower()
    if text in ("mine", "my", "ours", "local", "我的"):
        return DECISION_CONFLICT_MINE
    if text in ("empty", "blank", "none", "null", "空", "留空"):
        return DECISION_CONFLICT_EMPTY
    return DECISION_CONFLICT_THEIRS

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
YELLOW_FILL = PatternFill(start_color="FFFFDD", end_color="FFFFDD", fill_type="solid")
RED_FILL = PatternFill(start_color="FFDDDD", end_color="FFDDDD", fill_type="solid")
GREEN_FILL = PatternFill(start_color="DDFFDD", end_color="DDFFDD", fill_type="solid")
GRAY_FILL = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
LABEL_FONT = Font(bold=True, color="333333")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Row identity
# ---------------------------------------------------------------------------

def _get_rowname(ws, row_idx):
    """Get RowName from first column (A) of a data row."""
    val = ws.cell(row=row_idx, column=1).value
    if val is None:
        return ""
    return str(val).strip()


def _read_data_rows(wb, sheet_name, data_only=True):
    """
    Read data rows from a data sheet.
    Returns OrderedDict: {rowname: {col: (value, formula)}}
    """
    if sheet_name not in wb.sheetnames:
        return OrderedDict()

    ws = wb[sheet_name]
    rows = OrderedDict()

    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        rn = _get_rowname(ws, row_idx)
        if not rn:
            continue  # skip empty RowName rows
        row_data = {}
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            formula = None
            val = cell.value
            # Try to get formula — but with data_only=True, we can't.
            # We'll handle this at a higher level.
            row_data[col_idx] = val
        rows[rn] = row_data

    return rows


def _open_wb_any_ext(filepath, data_only=False, read_only=False):
    """Load workbook with any extension (.mine, .r4523 etc.) via temp copy."""
    import tempfile
    actual_path = filepath
    tmp_file = None
    if not str(filepath).lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        tmp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp_file.close()
        shutil.copy2(filepath, tmp_file.name)
        actual_path = tmp_file.name
    try:
        wb = openpyxl.load_workbook(actual_path, data_only=data_only, read_only=read_only)
    except Exception:
        if tmp_file:
            try:
                os.unlink(tmp_file.name)
            except Exception:
                pass
        raise
    if tmp_file:
        if read_only:
            # read_only workbooks read lazily from the file on disk; the temp
            # copy must survive until the workbook is closed.
            _orig_close = wb.close

            def _close_and_unlink():
                try:
                    _orig_close()
                finally:
                    try:
                        os.unlink(tmp_file.name)
                    except Exception:
                        pass

            wb.close = _close_and_unlink
        else:
            # Eager load: the file is fully parsed and already released.
            try:
                os.unlink(tmp_file.name)
            except Exception:
                pass
    return wb


def _external_link_index_map(filepath):
    """Map raw OOXML external reference numbers ([1], [2]) to workbook stems."""
    result = {}
    tmp_file = None
    try:
        actual_path = filepath
        if not str(filepath).lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
            tmp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            tmp_file.close()
            shutil.copy2(filepath, tmp_file.name)
            actual_path = tmp_file.name
        with zipfile.ZipFile(actual_path, "r") as z:
            workbook_xml = z.read("xl/workbook.xml").decode("utf-8", errors="replace")
            workbook_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
            ext_ref_ids = re.findall(r"<externalReference\b[^>]*\br:id=\"([^\"]+)\"", workbook_xml)
            rel_targets = {
                rid: target
                for rid, target in re.findall(
                    r"<Relationship\b[^>]*\bId=\"([^\"]+)\"[^>]*\bTarget=\"([^\"]+)\"",
                    workbook_rels,
                )
            }
            names = set(z.namelist())
            for idx, rid in enumerate(ext_ref_ids, start=1):
                target = rel_targets.get(rid)
                if not target:
                    continue
                link_xml = target.lstrip("/")
                if not link_xml.startswith("xl/"):
                    link_xml = f"xl/{link_xml}"
                rels_name = str(Path(link_xml).parent / "_rels" / (Path(link_xml).name + ".rels")).replace("\\", "/")
                if rels_name not in names:
                    continue
                rels_xml = z.read(rels_name).decode("utf-8", errors="replace")
                link_targets = re.findall(
                    r"<Relationship\b[^>]*\bType=\"[^\"]*externalLinkPath[^\"]*\"[^>]*\bTarget=\"([^\"]+)\"",
                    rels_xml,
                )
                if not link_targets:
                    continue
                chosen = next((t for t in link_targets if not t.lower().startswith("file:")), link_targets[0])
                name = os.path.basename(chosen.replace("\\", "/"))
                if name:
                    result[str(idx)] = os.path.splitext(name)[0]
    except Exception:
        return {}
    finally:
        if tmp_file:
            try:
                os.unlink(tmp_file.name)
            except Exception:
                pass
    return result


_EXTERNAL_REF_RE = re.compile(r"\[(\d+)\]")
_PREFERRED_EXTERNAL_LINK_DRIVE = os.environ.get("TABLE_MANAGER_LINK_DRIVE", "A")


def _normalize_external_formula_refs(formula, index_map):
    """Replace raw external slots with stable workbook names for preview/compare."""
    if not formula or not isinstance(formula, str) or not index_map:
        return formula

    def repl(match):
        key = match.group(1)
        return f"[{index_map.get(key, key)}]"

    return _EXTERNAL_REF_RE.sub(repl, formula)


def _preferred_raw_link_target(raw_relative_path, preferred_drive=None):
    drive = str(preferred_drive or _PREFERRED_EXTERNAL_LINK_DRIVE or "A").strip()
    drive = drive.rstrip(":\\/").upper() or "A"
    rel = str(raw_relative_path).replace("/", "\\").lstrip("\\")
    return f"file:///{drive}:\\DataTable\\Raw\\{rel}"


def _target_to_raw_relative_path(target, workbook_path):
    if not target:
        return None
    text = unquote(str(target).strip())
    lower = text.lower()
    if lower.startswith("file:///"):
        text = text[8:]
    elif lower.startswith("file://"):
        text = text[7:]

    normalized = text.replace("/", "\\")
    lower_normalized = normalized.lower()
    marker = "\\datatable\\raw\\"
    idx = lower_normalized.find(marker)
    if idx >= 0:
        return normalized[idx + len(marker):].lstrip("\\")
    marker_no_lead = "datatable\\raw\\"
    if lower_normalized.startswith(marker_no_lead):
        return normalized[len(marker_no_lead):].lstrip("\\")

    if re.match(r"^[A-Za-z]:\\", normalized) or normalized.startswith("\\\\"):
        return None

    workbook = Path(workbook_path)
    candidates = [
        workbook.parent / normalized,
        RAW_DIR / normalized,
    ]
    for candidate in candidates:
        try:
            return str(candidate.resolve().relative_to(RAW_DIR.resolve())).replace("/", "\\")
        except (OSError, ValueError):
            continue
    return None


def normalize_external_link_targets_to_preferred_drive(filepath, preferred_drive=None):
    """Rewrite external-link relationship targets to the shared Raw path.

    Excel may save links using the current user's working drive (for example
    D:). The final workbook should prefer A:\\DataTable\\Raw so every user's
    link targets are stable.
    """
    path = Path(filepath)
    if not path.is_file():
        return {"rewritten": 0, "drive": preferred_drive or _PREFERRED_EXTERNAL_LINK_DRIVE}

    rels_re = re.compile(r'(\bTarget=")([^"]*)(")')
    rewritten = 0
    tmp_path = path.with_name(f".{path.name}.normalize_links.tmp")

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("xl/externalLinks/_rels/") and info.filename.endswith(".rels"):
                text = data.decode("utf-8", errors="replace")

                def repl_target(match):
                    nonlocal rewritten
                    target = match.group(2)
                    rel = _target_to_raw_relative_path(target, path)
                    if not rel:
                        return match.group(0)
                    new_target = _preferred_raw_link_target(rel, preferred_drive)
                    if new_target != target:
                        rewritten += 1
                    return match.group(1) + new_target + match.group(3)

                text = rels_re.sub(repl_target, text)
                data = text.encode("utf-8")
            zout.writestr(info, data)

    os.replace(tmp_path, path)
    return {"rewritten": rewritten, "drive": preferred_drive or _PREFERRED_EXTERNAL_LINK_DRIVE}


def _denormalize_external_formula_refs(formula, name_map):
    """Replace preview workbook names with this workbook's external slots.

    Preview/compare uses stable names such as [CardTable]. Excel workbook XML
    formulas must still use its own external link slots such as [4].
    """
    if not formula or not isinstance(formula, str) or not name_map:
        return formula

    result = formula
    for name, idx in sorted(name_map.items(), key=lambda item: len(item[0]), reverse=True):
        result = result.replace(f"[{name}]", f"[{idx}]")
    return result


def rewrite_external_formula_refs_to_names(filepath, index_map=None):
    """Rewrite worksheet formula XML from [1]/[2] slots to workbook names.

    This is only safe for diagnostic/preview files. Final Excel workbooks must
    keep slot references because the OOXML external link table defines them.
    """
    path = Path(filepath)
    mapping = index_map or _external_link_index_map(path)
    if not mapping:
        return {"rewritten": 0, "mapping": mapping}

    formula_re = re.compile(r"(<f\b[^>]*>)(.*?)(</f>)", re.DOTALL)
    rewritten = 0
    tmp_path = path.with_name(f".{path.name}.rewrite_links.tmp")

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml"):
                text = data.decode("utf-8", errors="replace")

                def repl_formula(match):
                    nonlocal rewritten
                    body = match.group(2)
                    new_body = _normalize_external_formula_refs(body, mapping)
                    if new_body != body:
                        rewritten += 1
                    return match.group(1) + new_body + match.group(3)

                text = formula_re.sub(repl_formula, text)
                data = text.encode("utf-8")
            zout.writestr(info, data)

    os.replace(tmp_path, path)
    return {"rewritten": rewritten, "mapping": mapping}


def _read_all_rows_with_formulas(filepath, include_origins=False):
    """
    Read workbook twice: once for values, once for formulas.
    Uses read_only=True for much faster loading on large files.
    Handles non-.xlsx extensions (.mine, .r4523 etc.) by copying to temp file.
    Returns: {sheet_name: {rowname: {col: (value, formula)}}}
    When ``include_origins`` is true, also returns the original worksheet
    coordinate for every cell.  The latter is used only to normalize formulas
    for comparison; formulas kept for preview and resolve remain unchanged.
    """
    import tempfile
    import shutil

    actual_path = filepath
    tmp_file = None

    # openpyxl rejects non-standard extensions like .r4523 or .mine
    if not str(filepath).lower().endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        tmp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp_file.close()
        shutil.copy2(filepath, tmp_file.name)
        actual_path = tmp_file.name

    # NOTE: the temp copy must survive until both read_only workbooks are
    # closed (cleanup happens in the finally block at the end of this function).
    try:
        wb_val = openpyxl.load_workbook(actual_path, data_only=True, read_only=True)
    except Exception as e:
        print(f"[WARN] Cannot open (values) {filepath}: {e}")
        if tmp_file:
            try:
                os.unlink(tmp_file.name)
            except Exception:
                pass
        return ({}, {}, {}) if include_origins else ({}, {})

    try:
        wb_fml = openpyxl.load_workbook(actual_path, data_only=False, read_only=True)
    except Exception as e:
        print(f"[WARN] Cannot open (formulas) {filepath}: {e}")
        wb_fml = None

    result = {}
    headers = {}
    origins = {}
    try:
        for sn in wb_val.sheetnames:
            if not sn.startswith("#"):
                continue
            ws_val = wb_val[sn]
            sheet_rows = OrderedDict()
            sheet_origins = OrderedDict()

            # Capture header rows 1-3 during same open (avoids separate _read_headers call)
            hdr_rows = []
            for hdr_row in ws_val.iter_rows(min_row=1, max_row=3, values_only=True):
                hdr_rows.append(tuple(hdr_row))
            max_hdr_col = max((len(r) for r in hdr_rows), default=1)
            headers[sn] = {
                "row1": hdr_rows[0] if len(hdr_rows) > 0 else (),
                "row2": hdr_rows[1] if len(hdr_rows) > 1 else (),
                "row3": hdr_rows[2] if len(hdr_rows) > 2 else (),
                "max_col": max_hdr_col,
            }

            # read_only mode: collect rows by index from both loads
            val_by_idx = {}
            for row_idx, row in enumerate(ws_val.iter_rows(min_row=DATA_START_ROW, values_only=True), start=DATA_START_ROW):
                val_by_idx[row_idx] = row

            fml_by_idx = {}
            if wb_fml and sn in wb_fml.sheetnames:
                ws_fml = wb_fml[sn]
                for row_idx, row in enumerate(ws_fml.iter_rows(min_row=DATA_START_ROW, values_only=True), start=DATA_START_ROW):
                    fml_by_idx[row_idx] = row

            # Determine RowName for each row:
            # Priority: value row col[0] → formula row col[0] (as formula text)
            # If both are None/empty, try col[1] value as fallback (since many tables use =B4 for RowName)
            all_indices = sorted(set(val_by_idx.keys()) | set(fml_by_idx.keys()))
            for row_idx in all_indices:
                val_row = val_by_idx.get(row_idx, ())
                fml_row = fml_by_idx.get(row_idx, ())

                # Determine RowName
                rn = ""
                # Try value first (actual computed value)
                if val_row and val_row[0] is not None:
                    rn = str(val_row[0]).strip()
                # If value is None (formula without cache), use col B value as RowName
                # (common pattern: col A = =B4, meaning RowName IS the value in col B)
                if not rn and val_row and len(val_row) > 1 and val_row[1] is not None:
                    rn = str(val_row[1]).strip()
                # Last resort: formula text from col A (e.g. "=B4")
                if not rn and fml_row and fml_row[0] is not None:
                    fml_rn = str(fml_row[0]).strip()
                    if fml_rn.startswith("="):
                        # Use col B value from formula row if available
                        if fml_row and len(fml_row) > 1 and fml_row[1] is not None:
                            rn = str(fml_row[1]).strip()
                        else:
                            rn = fml_rn
                    else:
                        rn = fml_rn

                if not rn:
                    continue

                row_data = {}
                row_origins = {}
                max_col = max(len(val_row) if val_row else 0, len(fml_row) if fml_row else 0)

                for col_idx in range(max_col):
                    val = val_row[col_idx] if val_row and col_idx < len(val_row) else None
                    fml = None
                    if fml_row and col_idx < len(fml_row):
                        fml_cell = fml_row[col_idx]
                        if isinstance(fml_cell, str) and fml_cell.startswith("="):
                            fml = fml_cell
                        elif isinstance(fml_cell, ArrayFormula):
                            # Unwrap array formulas so the merge pipeline
                            # treats them as formulas instead of flattening
                            # them to their cached value.
                            fml = fml_cell.text
                    row_data[col_idx + 1] = (val, fml)
                    row_origins[col_idx + 1] = (row_idx, col_idx + 1)
                if rn in sheet_rows:
                    print(f"[WARN] Duplicate RowName '{rn}' in sheet '{sn}' of {filepath}; later row overwrites earlier one.")
                sheet_rows[rn] = row_data
                sheet_origins[rn] = row_origins

            result[sn] = sheet_rows
            origins[sn] = sheet_origins
    finally:
        try:
            wb_val.close()
        except Exception:
            pass
        if wb_fml:
            try:
                wb_fml.close()
            except Exception:
                pass
        # read_only workbooks read lazily - delete the temp copy only now.
        if tmp_file:
            try:
                os.unlink(tmp_file.name)
            except Exception:
                pass
    if include_origins:
        return result, headers, origins
    return result, headers


def _norm_header_value(value):
    """Normalize a header cell for stable column identity matching."""
    if value is None:
        return ""
    return str(value).strip()


def _header_signatures(header_info):
    """Build physical-column -> semantic-column signature mapping.

    English field names are the preferred identity. Some localization tables
    contain repeated placeholder English names such as "None"; in that case
    the full visible header tuple is used. Identical header tuples (duplicate
    Chinese names) are disambiguated by an occurrence index so every physical
    column keeps its own canonical column — collapsing duplicates would drop
    a column and shift all later columns by one.
    """
    row1 = list(header_info.get("row1", ()) or ())
    row2 = list(header_info.get("row2", ()) or ())
    row3 = list(header_info.get("row3", ()) or ())
    max_col = max(header_info.get("max_col", 0), len(row1), len(row2), len(row3))

    placeholder_en = {"none", "null", "nan", "-"}
    en_values = []
    for idx in range(max_col):
        en = _norm_header_value(row2[idx] if idx < len(row2) else "")
        if en and en.lower() not in placeholder_en:
            en_values.append(en)
    en_counts = {}
    for en in en_values:
        en_counts[en] = en_counts.get(en, 0) + 1

    result = OrderedDict()
    header_occurrences = {}
    for col in range(1, max_col + 1):
        cn = _norm_header_value(row1[col - 1] if col - 1 < len(row1) else "")
        en = _norm_header_value(row2[col - 1] if col - 1 < len(row2) else "")
        typ = _norm_header_value(row3[col - 1] if col - 1 < len(row3) else "")
        if en and en.lower() not in placeholder_en and en_counts.get(en) == 1:
            sig = ("en", en)
        elif cn or en or typ:
            # Duplicate header tuples (e.g. two columns both named 卡牌标签（3）)
            # must stay distinct — collapsing them into one signature drops a
            # canonical column and shifts every later column by one, which
            # turns relative formula references into self-references.
            base_sig = ("header", cn, en, typ)
            header_occurrences[base_sig] = header_occurrences.get(base_sig, 0) + 1
            sig = (*base_sig, header_occurrences[base_sig])
        else:
            sig = ("blank", col)
        result[col] = {
            "sig": sig,
            "row1": cn,
            "row2": en,
            "row3": typ,
            "source_col": col,
        }
    return result


def _canonical_columns(*header_infos):
    """Merge header definitions into a canonical column order.

    The caller should pass the preferred output layout first. For merges this
    is mine, then theirs, then base; for diffs this is new, then old.
    """
    columns = []
    seen = set()
    for header_info in header_infos:
        for meta in _header_signatures(header_info or {}).values():
            sig = meta["sig"]
            if sig in seen:
                continue
            seen.add(sig)
            columns.append({
                "sig": sig,
                "row1": meta["row1"],
                "row2": meta["row2"],
                "row3": meta["row3"],
            })
    if not columns:
        columns.append({"sig": ("blank", 1), "row1": "", "row2": "", "row3": ""})
    return columns


def _remap_rows_to_columns(rows, header_info, columns):
    """Remap physical column-indexed row data to canonical header columns."""
    sig_to_canonical = {col["sig"]: idx + 1 for idx, col in enumerate(columns)}
    physical_to_sig = {phys_col: meta["sig"] for phys_col, meta in _header_signatures(header_info or {}).items()}

    remapped = OrderedDict()
    for rowname, row_data in rows.items():
        new_row = {}
        for phys_col, cell in row_data.items():
            sig = physical_to_sig.get(phys_col)
            if sig is None:
                continue
            canonical_col = sig_to_canonical.get(sig)
            if canonical_col is not None:
                new_row[canonical_col] = cell
        remapped[rowname] = new_row
    return remapped


def _remap_origins_to_columns(origins, header_info, columns):
    """Remap source cell coordinates to the canonical comparison columns."""
    sig_to_canonical = {col["sig"]: idx + 1 for idx, col in enumerate(columns)}
    physical_to_sig = {
        phys_col: meta["sig"]
        for phys_col, meta in _header_signatures(header_info or {}).items()
    }

    remapped = OrderedDict()
    for rowname, row_origins in origins.items():
        new_origins = {}
        for phys_col, origin in row_origins.items():
            canonical_col = sig_to_canonical.get(physical_to_sig.get(phys_col))
            if canonical_col is not None:
                new_origins[canonical_col] = origin
        remapped[rowname] = new_origins
    return remapped


def _columns_to_header_info(columns):
    """Convert canonical column metadata back to the header-info shape."""
    return {
        "row1": tuple(col.get("row1", "") for col in columns),
        "row2": tuple(col.get("row2", "") for col in columns),
        "row3": tuple(col.get("row3", "") for col in columns),
        "max_col": len(columns),
    }


def align_datasets_by_headers(preferred_data, preferred_headers, other_data, other_headers):
    """Align two workbook datasets by semantic headers, preserving preferred order."""
    aligned_preferred = {}
    aligned_other = {}
    aligned_headers = {}
    all_sheets = set(preferred_data.keys()) | set(other_data.keys())
    for sn in all_sheets:
        columns = _canonical_columns(preferred_headers.get(sn, {}), other_headers.get(sn, {}))
        aligned_preferred[sn] = _remap_rows_to_columns(
            preferred_data.get(sn, OrderedDict()),
            preferred_headers.get(sn, {}),
            columns,
        )
        aligned_other[sn] = _remap_rows_to_columns(
            other_data.get(sn, OrderedDict()),
            other_headers.get(sn, {}),
            columns,
        )
        aligned_headers[sn] = _columns_to_header_info(columns)
    return aligned_preferred, aligned_other, aligned_headers


# ---------------------------------------------------------------------------
# Key normalization — shared by preview parsing and resolve_merge fallbacks
# ---------------------------------------------------------------------------

def _strip_col(val):
    """Extract the computed VALUE from a display-format cell.
    '=FORMULA → VALUE' → 'VALUE' (key matching by computed value)
    '=FORMULA'        → 'FORMULA' (fallback: no cached value)
    plain text        → text as-is
    """
    s = str(val).strip() if val is not None else ""
    if s.startswith("'="):
        s = s[1:]
    if chr(0x2192) in s:
        parts = s.split(f" {chr(0x2192)} ", 1)
        if len(parts) > 1 and parts[1]:
            return parts[1]  # computed value (right side of →)
        return parts[0]  # formula only, no value
    return s


# ---------------------------------------------------------------------------
# Conflict detection
# ---------------------------------------------------------------------------

class ConflictDetector:
    def __init__(self, base_path, mine_path, theirs_path):
        self.base_path = base_path
        self.mine_path = mine_path
        self.theirs_path = theirs_path
        self.base_external_refs = _external_link_index_map(base_path)
        self.mine_external_refs = _external_link_index_map(mine_path)
        self.theirs_external_refs = _external_link_index_map(theirs_path)
        self.base_data, self.base_headers, self.base_origins = _read_all_rows_with_formulas(
            base_path, include_origins=True
        )
        self.mine_data, self.mine_headers, self.mine_origins = _read_all_rows_with_formulas(
            mine_path, include_origins=True
        )
        self.theirs_data, self.theirs_headers, self.theirs_origins = _read_all_rows_with_formulas(
            theirs_path, include_origins=True
        )

    def detect(self):
        """
        Run 3-way diff on all data sheets.
        Returns: {
            sheet_name: {
                "headers": {col: (cn_name, en_name, type_val)},
                "rows": [RowDiff, ...]
            }
        }
        """
        all_sheets = set(self.base_data.keys()) | set(self.mine_data.keys()) | set(self.theirs_data.keys())
        result = {}

        for sn in sorted(all_sheets):
            base_rows = self.base_data.get(sn, OrderedDict())
            mine_rows = self.mine_data.get(sn, OrderedDict())
            theirs_rows = self.theirs_data.get(sn, OrderedDict())
            columns = _canonical_columns(
                self.mine_headers.get(sn, {}),
                self.theirs_headers.get(sn, {}),
                self.base_headers.get(sn, {}),
            )
            base_rows = _remap_rows_to_columns(base_rows, self.base_headers.get(sn, {}), columns)
            mine_rows = _remap_rows_to_columns(mine_rows, self.mine_headers.get(sn, {}), columns)
            theirs_rows = _remap_rows_to_columns(theirs_rows, self.theirs_headers.get(sn, {}), columns)
            base_origins = _remap_origins_to_columns(
                self.base_origins.get(sn, OrderedDict()), self.base_headers.get(sn, {}), columns
            )
            mine_origins = _remap_origins_to_columns(
                self.mine_origins.get(sn, OrderedDict()), self.mine_headers.get(sn, {}), columns
            )
            theirs_origins = _remap_origins_to_columns(
                self.theirs_origins.get(sn, OrderedDict()), self.theirs_headers.get(sn, {}), columns
            )

            all_keys = list(OrderedDict.fromkeys(
                list(base_rows.keys()) + list(mine_rows.keys()) + list(theirs_rows.keys())
            ))

            row_diffs = []
            for key in all_keys:
                b = base_rows.get(key)
                m = mine_rows.get(key)
                t = theirs_rows.get(key)

                rd = self._diff_row(
                    key, b, m, t,
                    base_origins.get(key), mine_origins.get(key), theirs_origins.get(key),
                )
                row_diffs.append(rd)

            result[sn] = {
                "headers": _columns_to_header_info(columns),
                "columns": columns,
                "rows": row_diffs,
            }

        return result

    def _diff_row(
        self, rowname, base_row, mine_row, theirs_row,
        base_origins=None, mine_origins=None, theirs_origins=None,
    ):
        """
        Diff a single row across three versions.
        base_row, mine_row, theirs_row: {col: (value, formula)} or None
        """
        # Determine row status
        m_exists = mine_row is not None
        t_exists = theirs_row is not None
        b_exists = base_row is not None

        # Gather all columns
        all_cols = set()
        for r in [base_row, mine_row, theirs_row]:
            if r:
                all_cols.update(r.keys())
        all_cols = sorted(all_cols)

        cells = []
        has_conflict = False
        # Per-side "kept the row but changed it vs base" flags, used for
        # delete-vs-modify conflict detection below.
        mine_row_modified = False
        theirs_row_modified = False

        for col in all_cols:
            bv = base_row.get(col, (None, None)) if b_exists else (None, None)
            mv = mine_row.get(col, (None, None)) if m_exists else (None, None)
            tv = theirs_row.get(col, (None, None)) if t_exists else (None, None)

            b_val, b_fml = bv
            m_val, m_fml = mv
            t_val, t_fml = tv
            b_fml_display = _normalize_external_formula_refs(b_fml, self.base_external_refs)
            m_fml_display = _normalize_external_formula_refs(m_fml, self.mine_external_refs)
            t_fml_display = _normalize_external_formula_refs(t_fml, self.theirs_external_refs)

            def compare_formula(formula, origins):
                origin = (origins or {}).get(col)
                if not origin:
                    return formula
                return normalize_formula_for_compare(formula, origin[0], origin[1])

            b_fml_compare = compare_formula(b_fml_display, base_origins)
            m_fml_compare = compare_formula(m_fml_display, mine_origins)
            t_fml_compare = compare_formula(t_fml_display, theirs_origins)

            # Prefer cached values.  If a file was saved without those caches,
            # compare a location-independent formula signature instead of raw
            # A1 text so inserted rows do not become false modifications.
            b_v = str(b_val) if b_val is not None else (b_fml_compare or "")
            m_v = str(m_val) if m_val is not None else (m_fml_compare or "")
            t_v = str(t_val) if t_val is not None else (t_fml_compare or "")

            if m_v == t_v:
                cell_status = SAME
            elif not b_exists and m_exists and t_exists:
                # Both added but different → conflict
                cell_status = CONFLICT
                has_conflict = True
            elif not b_exists and m_exists and not t_exists:
                cell_status = MINE_ADDED
            elif not b_exists and not m_exists and t_exists:
                cell_status = THEIRS_ADDED
            elif b_exists and not m_exists and t_exists:
                cell_status = THEIRS_MOD
                if b_v != t_v:
                    theirs_row_modified = True
            elif b_exists and m_exists and not t_exists:
                cell_status = MINE_MOD
                if b_v != m_v:
                    mine_row_modified = True
            elif m_v == b_v:
                cell_status = THEIRS_MOD
            elif t_v == b_v:
                cell_status = MINE_MOD
            else:
                cell_status = CONFLICT
                has_conflict = True

            cells.append({
                "col": col,
                "status": cell_status,
                "base_value": b_val,
                "base_formula": b_fml,
                "base_formula_display": b_fml_display,
                "mine_value": m_val,
                "mine_formula": m_fml,
                "mine_formula_display": m_fml_display,
                "theirs_value": t_val,
                "theirs_formula": t_fml,
                "theirs_formula_display": t_fml_display,
            })

        # Overall row status
        if not b_exists and m_exists and not t_exists:
            row_status = MINE_ADDED
        elif not b_exists and not m_exists and t_exists:
            row_status = THEIRS_ADDED
        elif not b_exists and m_exists and t_exists:
            row_status = BOTH_ADDED
        elif b_exists and not m_exists and t_exists:
            # Mine deleted the row. Keep the historical default decision
            # [删除] even when theirs modified it (integration fixtures rely
            # on it); theirs' version is still shown in the preview via
            # "other_side_modified" so the change is not lost silently.
            row_status = MINE_DELETED
        elif b_exists and m_exists and not t_exists:
            # Theirs deleted the row; if mine also modified it, this is a
            # delete-vs-modify conflict - never silently drop mine's changes.
            if mine_row_modified:
                row_status = CONFLICT
                has_conflict = True
            else:
                row_status = THEIRS_DELETED
        elif has_conflict:
            row_status = CONFLICT
        elif any(c["status"] == MINE_MOD for c in cells):
            row_status = MINE_MOD
        elif any(c["status"] == THEIRS_MOD for c in cells):
            row_status = THEIRS_MOD
        else:
            row_status = SAME

        return {
            "rowname": rowname,
            "status": row_status,
            "has_conflict": has_conflict,
            "cells": cells,
            # True when mine deleted the row but theirs modified it: the
            # preview must still write the [仓库] source row so theirs'
            # change stays visible next to the [删除] default decision.
            "other_side_modified": (
                row_status == MINE_DELETED and theirs_row_modified
            ),
        }


# ---------------------------------------------------------------------------
# Merge preview generator
# ---------------------------------------------------------------------------

class MergePreview:
    """Generate merge preview Excel file with [我的]/[仓库]/[最终] rows."""

    def __init__(self, base_path, mine_path, theirs_path, output_path, conflict_default=DECISION_CONFLICT_THEIRS):
        self.base_path = base_path
        self.mine_path = mine_path
        self.theirs_path = theirs_path
        self.output_path = output_path
        self.conflict_default = _normalize_conflict_default(conflict_default)
        self.detector = ConflictDetector(base_path, mine_path, theirs_path)

    @staticmethod
    def _read_headers(filepath):
        """Read header rows (1-3) and max_col per sheet."""
        try:
            wb = _open_wb_any_ext(filepath, data_only=True, read_only=True)
        except Exception:
            return {}
        try:
            result = {}
            for sn in wb.sheetnames:
                if not sn.startswith("#"):
                    continue
                ws = wb[sn]
                rows = []
                for i, row in enumerate(ws.iter_rows(min_row=1, max_row=3, values_only=True)):
                    rows.append(row)
                max_col = max((len(r) for r in rows), default=1)
                result[sn] = {
                    "row1": rows[0] if len(rows) > 0 else (),
                    "row2": rows[1] if len(rows) > 1 else (),
                    "row3": rows[2] if len(rows) > 2 else (),
                    "max_col": max_col,
                }
            return result
        finally:
            wb.close()

    def generate(self):
        """Generate the merge preview workbook."""
        diffs = self.detector.detect()

        out_wb = openpyxl.Workbook()
        # Remove default sheet (only after we know we have data to write)
        first_sheet = True

        if not diffs:
            # Keep default sheet visible with a message
            ws = out_wb.active
            ws.title = "NoData"
            ws.cell(row=1, column=1, value="No data sheets (#-prefixed) found in conflicted files.")
            out_wb.save(self.output_path)
            return self.output_path

        for sn in sorted(diffs.keys()):
            sheet_diffs = diffs[sn]
            if first_sheet:
                ws = out_wb.active
                ws.title = sn
                first_sheet = False
            else:
                ws = out_wb.create_sheet(title=sn)

            canonical_hdr = sheet_diffs.get("headers", {})
            max_col = max(1, canonical_hdr.get("max_col", 1))

            # Row 1: Chinese headers — col A=[标记], col B=[修改类型], col C+ = data
            ws.cell(row=1, column=1, value="[标记]")
            ws.cell(row=1, column=1).fill = HEADER_FILL
            ws.cell(row=1, column=1).font = HEADER_FONT
            ws.cell(row=1, column=2, value="[修改类型]")
            ws.cell(row=1, column=2).fill = HEADER_FILL
            ws.cell(row=1, column=2).font = HEADER_FONT

            canonical_r1 = canonical_hdr.get("row1", ())
            for col in range(3, max_col + 3):
                data_col = col - 2
                header_val = canonical_r1[data_col - 1] if data_col <= len(canonical_r1) else ""
                cell = ws.cell(row=1, column=col, value=header_val)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT

            # Row 2: English field names
            ws.cell(row=2, column=1, value="(EN)")
            ws.cell(row=2, column=1).fill = HEADER_FILL
            ws.cell(row=2, column=1).font = HEADER_FONT
            ws.cell(row=2, column=2, value="ChangeType")
            ws.cell(row=2, column=2).fill = HEADER_FILL
            ws.cell(row=2, column=2).font = HEADER_FONT
            canonical_r2 = canonical_hdr.get("row2", ())
            for col in range(3, max_col + 3):
                data_col = col - 2
                en_val = canonical_r2[data_col - 1] if data_col <= len(canonical_r2) else ""
                cell = ws.cell(row=2, column=col, value=en_val)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT

            # Row 3: Type info
            ws.cell(row=3, column=1, value="(type)")
            ws.cell(row=3, column=1).fill = HEADER_FILL
            ws.cell(row=3, column=1).font = HEADER_FONT
            ws.cell(row=3, column=2, value="")
            ws.cell(row=3, column=2).fill = HEADER_FILL
            canonical_r3 = canonical_hdr.get("row3", ())
            for col in range(3, max_col + 3):
                data_col = col - 2
                type_val = canonical_r3[data_col - 1] if data_col <= len(canonical_r3) else ""
                cell = ws.cell(row=3, column=col, value=type_val)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT

            # Write data rows — skip SAME rows, only show conflicts
            current_row = DATA_START_ROW  # row 4
            final_modify_cells = []  # [修改]/[删除] dropdowns
            final_conflict_cells = []  # Mine/theirs/空 dropdowns

            # Map row status to change type labels and default decision
            _CT_MAP = {
                MINE_MOD:     ("[已修改]", None),
                THEIRS_MOD:   (None, "[已修改]"),
                CONFLICT:     ("[已修改]", "[已修改]"),
                MINE_ADDED:   ("[已添加]", None),
                THEIRS_ADDED: (None, "[已添加]"),
                BOTH_ADDED:   ("[已添加]", "[已添加]"),
                MINE_DELETED: ("[已删除]", None),
                THEIRS_DELETED: (None, "[已删除]"),
            }
            # Delete-type statuses → default decision is [删除]; all else → [修改]
            _DELETE_STATUSES = {MINE_DELETED, THEIRS_DELETED}

            for rd in sheet_diffs["rows"]:
                status = rd["status"]

                # Skip SAME rows (no conflict)
                if status == SAME:
                    continue

                mine_label, theirs_label = _CT_MAP.get(status, (None, None))
                if status == MINE_DELETED and rd.get("other_side_modified"):
                    # Mine deleted this row but theirs modified it - show
                    # theirs' version as well so the change stays visible.
                    theirs_label = "[已修改]"
                default_decision = DECISION_DELETE if status in _DELETE_STATUSES else DECISION_MODIFY
                if status == CONFLICT:
                    default_decision = self.conflict_default

                has_mine = self._get_row_data(rd, "mine")
                has_theirs = self._get_row_data(rd, "theirs")

                # Determine display side for [我的] row:
                # - MINE_DELETED: mine has no data → use base as display side
                # - Others: use mine
                mine_display_side = "mine"
                if status == MINE_DELETED:
                    mine_display_side = "base"
                elif not has_mine and mine_label:
                    mine_display_side = "base"

                # Determine display side for [仓库] row:
                theirs_display_side = "theirs"
                if status == THEIRS_DELETED:
                    theirs_display_side = "base"
                elif not has_theirs and theirs_label:
                    theirs_display_side = "base"

                # Write [我的] row if this side has a change label
                if mine_label:
                    self._write_labeled_row(
                        ws, current_row, max_col, "[我的]", rd, mine_display_side,
                        YELLOW_FILL if status in (MINE_MOD, MINE_ADDED) else None,
                        change_type_label=mine_label,
                    )
                    current_row += 1

                # Write [仓库] row if this side has a change label
                if theirs_label:
                    self._write_labeled_row(
                        ws, current_row, max_col, "[仓库]", rd, theirs_display_side,
                        YELLOW_FILL if status in (THEIRS_MOD, THEIRS_ADDED) else None,
                        change_type_label=theirs_label,
                    )
                    current_row += 1

                # Write [最终] row — always, for all non-SAME rows
                fill_side = "mine"
                if status in (THEIRS_ADDED, THEIRS_MOD):
                    fill_side = "theirs"
                if status == CONFLICT:
                    fill_side = "theirs"
                if status == MINE_DELETED:
                    fill_side = "base"

                self._write_final_preview_row(
                    ws, current_row, max_col, rd, fill_side,
                    default_decision=default_decision,
                )
                if status == CONFLICT:
                    final_conflict_cells.append(f"B{current_row}")
                else:
                    final_modify_cells.append(f"B{current_row}")
                current_row += 1

                # Blank separator row
                current_row += 1

            # Add DataValidation for [最终] B column dropdowns
            if final_modify_cells:
                dv = DataValidation(type="list", formula1=f'"{DECISION_MODIFY},{DECISION_DELETE}"', allow_blank=True)
                dv.error = "请选择 [修改] 或 [删除]"
                dv.errorTitle = "无效选择"
                dv.prompt = "选择处理方式"
                dv.promptTitle = "修改类型"
                ws.add_data_validation(dv)
                for cell_ref in final_modify_cells:
                    dv.add(ws[cell_ref])

            if final_conflict_cells:
                dv = DataValidation(
                    type="list",
                    formula1=f'"{DECISION_CONFLICT_MINE},{DECISION_CONFLICT_THEIRS},{DECISION_CONFLICT_EMPTY}"',
                    allow_blank=True,
                )
                dv.error = "请选择 Mine、theirs 或 空"
                dv.errorTitle = "无效选择"
                dv.prompt = "选择冲突单元格优先使用哪一侧"
                dv.promptTitle = "冲突优先"
                ws.add_data_validation(dv)
                for cell_ref in final_conflict_cells:
                    dv.add(ws[cell_ref])

            # Adjust column widths
            ws.column_dimensions["A"].width = 10
            ws.column_dimensions["B"].width = 12
            for col in range(3, max_col + 3):
                ws.column_dimensions[get_column_letter(col)].width = 18

            # Freeze header
            ws.freeze_panes = f"A{DATA_START_ROW}"

        self._write_metadata_sheet(out_wb)
        out_wb.save(self.output_path)
        return self.output_path

    def _write_metadata_sheet(self, out_wb):
        """Record source workbook sheet order so resolve/debug does not rely on index guesses."""
        meta_name = "_merge_meta"
        if meta_name in out_wb.sheetnames:
            del out_wb[meta_name]
        ws = out_wb.create_sheet(meta_name)
        ws.sheet_state = "hidden"

        ws.append(["kind", "role", "path", "sheet_name", "sheet_index", "is_data_sheet"])
        for role, path in (
            ("base", self.base_path),
            ("mine", self.mine_path),
            ("theirs", self.theirs_path),
        ):
            if not path:
                continue
            try:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
            except Exception as exc:
                ws.append(["workbook_error", role, str(path), str(exc), None, None])
                continue
            try:
                for idx, sheet_name in enumerate(wb.sheetnames, start=1):
                    ws.append([
                        "sheet_order",
                        role,
                        str(path),
                        sheet_name,
                        idx,
                        str(sheet_name).startswith("#"),
                    ])
            finally:
                wb.close()

    def _get_row_data(self, rd, side):
        """Check if any cell has data for a side (mine/theirs). Returns bool."""
        for c in rd["cells"]:
            val = c.get(f"{side}_value")
            if val is not None and str(val).strip() != "":
                return True
            # Formula-only cells (no cached value) still count as data -
            # otherwise the preview falls back to showing base values.
            fml = c.get(f"{side}_formula")
            if fml is not None and str(fml).strip() != "":
                return True
        return False

    def _write_same_row(self, ws, row, max_col, rd):
        """Write a row for SAME status — no label, just data."""
        ws.cell(row=row, column=1, value="")

        for col_idx in range(1, max_col + 1):
            cell_data = next((c for c in rd["cells"] if c["col"] == col_idx), None)
            excel_col = col_idx + 1

            if cell_data:
                val = cell_data.get("mine_value")
                fml = cell_data.get("mine_formula_display") or cell_data.get("mine_formula")
                display = format_display(fml, val) if fml else (val if val is not None else "")
                ws.cell(row=row, column=excel_col, value=display)
            else:
                ws.cell(row=row, column=excel_col, value="")

    def _write_labeled_row(self, ws, row, max_col, label, rd, side, fill, change_type_label=None):
        """Write a [我的] or [仓库] row."""
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = LABEL_FONT
        if fill:
            label_cell.fill = fill

        # B column: change type annotation
        if change_type_label:
            ct_cell = ws.cell(row=row, column=2, value=change_type_label)
            ct_cell.font = LABEL_FONT
            ct_cell.fill = fill if fill else GRAY_FILL
        else:
            ws.cell(row=row, column=2, value="").fill = GRAY_FILL if not fill else fill

        for col_idx in range(1, max_col + 1):
            cell_data = next((c for c in rd["cells"] if c["col"] == col_idx), None)
            excel_col = col_idx + 2  # offset for label + change_type columns

            if cell_data:
                val = cell_data.get(f"{side}_value")
                fml = cell_data.get(f"{side}_formula_display") or cell_data.get(f"{side}_formula")
                display = format_display(fml, val) if fml else (val if val is not None else "")

                c = ws.cell(row=row, column=excel_col, value=display)

                # Highlight only cells that THIS side actually changed
                if cell_data["status"] == CONFLICT:
                    c.fill = RED_FILL
                elif side == "mine" and cell_data["status"] in (MINE_MOD, MINE_ADDED):
                    c.fill = fill if fill else YELLOW_FILL
                elif side == "theirs" and cell_data["status"] in (THEIRS_MOD, THEIRS_ADDED):
                    c.fill = fill if fill else YELLOW_FILL
                # SAME cells or changes by the other side → no highlight
            else:
                ws.cell(row=row, column=excel_col, value="")

        ws.cell(row=row, column=1).fill = GRAY_FILL if not fill else fill
        ws.cell(row=row, column=1).font = LABEL_FONT

    def _write_final_preview_row(self, ws, row, max_col, rd, fill_side, default_decision=DECISION_MODIFY):
        """Write the [最终] row, pre-filled from fill_side ('mine' or 'theirs')."""
        label = "[最终]"
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = LABEL_FONT
        label_cell.fill = GREEN_FILL

        # B column: decision dropdown, pre-fill default
        decision_cell = ws.cell(row=row, column=2, value=default_decision)
        decision_cell.font = LABEL_FONT
        decision_cell.fill = GREEN_FILL

        for col_idx in range(1, max_col + 1):
            cell_data = next((c for c in rd["cells"] if c["col"] == col_idx), None)
            excel_col = col_idx + 2  # offset for label + change_type columns

            if cell_data:
                cell_status = cell_data["status"]
                side = fill_side
                if cell_status == CONFLICT:
                    if default_decision == DECISION_CONFLICT_MINE:
                        side = "mine"
                    elif default_decision == DECISION_CONFLICT_EMPTY:
                        side = None
                    else:
                        side = "theirs"
                elif cell_status in (THEIRS_MOD, THEIRS_ADDED):
                    side = "theirs"
                elif cell_status in (MINE_MOD, MINE_ADDED):
                    side = "mine"

                # Pre-fill from selected side. For conflict rows this means
                # non-conflict cells keep their naturally resolved side while
                # conflict cells follow the B-column priority.
                if side is None:
                    val = ""
                    fml = None
                else:
                    val = cell_data.get(f"{side}_value")
                    fml = cell_data.get(f"{side}_formula_display") or cell_data.get(f"{side}_formula")
                display = format_display(fml, val) if fml else (val if val is not None else "")

                c = ws.cell(row=row, column=excel_col, value=display)
                c.fill = GREEN_FILL

                # Only highlight CONFLICT cells in red — user needs to decide these
                if cell_data["status"] == CONFLICT:
                    c.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            else:
                ws.cell(row=row, column=excel_col, value="").fill = GREEN_FILL


# ---------------------------------------------------------------------------
# External link preservation
# ---------------------------------------------------------------------------


def _read_external_link_cache(filepath):
    """
    Read and cache all external-link data from an xlsx zip before openpyxl
    can corrupt it on load/save. Returns None if no external links present.
    """
    if not os.path.isfile(filepath):
        return None

    import re

    with zipfile.ZipFile(filepath, 'r') as z:
        extlink_files = {}
        extlink_rel_blocks = []
        extlink_ct_blocks = []

        for n in z.namelist():
            if n.startswith("xl/externalLinks/") or n.startswith("externalLinks/"):
                extlink_files[n] = z.read(n)

        if not extlink_files:
            return None

        try:
            rels_str = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            for m in re.finditer(
                r'<Relationship\s[^>]*Type="[^"]*externalLink[^"]*"[^>]*/?>',
                rels_str, re.IGNORECASE
            ):
                extlink_rel_blocks.append(m.group(0))
        except KeyError:
            pass

        try:
            ct_str = z.read("[Content_Types].xml").decode("utf-8")
            for m in re.finditer(
                r'<Override\s[^>]*PartName="[^"]*externalLink[^"]*"[^>]*/?>',
                ct_str, re.IGNORECASE
            ):
                extlink_ct_blocks.append(m.group(0))
        except KeyError:
            pass

        return {
            "files": extlink_files,
            "rel_blocks": extlink_rel_blocks,
            "ct_blocks": extlink_ct_blocks,
        }


def _inject_dynamic_array_metadata(source_path, saved_path, fallback_path=None):
    """Restore dynamic-array metadata that openpyxl drops on save.

    Cells with modern functions (XLOOKUP etc.) inside CSE array formulas rely
    on xl/metadata.xml (XLDAPR) plus a per-cell cm="1" attribute. openpyxl
    discards both, so after a save Excel re-evaluates those array formulas
    with implicit intersection and returns wrong results (#VALUE! or just the
    first element). Re-add metadata.xml at zip level and tag every
    <f t="array"> cell with cm="1".
    """
    metadata_xml = None
    for candidate in (source_path, fallback_path):
        if not candidate or not os.path.isfile(candidate):
            continue
        try:
            with zipfile.ZipFile(candidate) as z:
                if "xl/metadata.xml" in z.namelist():
                    data = z.read("xl/metadata.xml")
                    if b"XLDAPR" in data:
                        metadata_xml = data
                        break
        except Exception:
            continue
    if metadata_xml is None or not os.path.isfile(saved_path):
        return

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(saved_path, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                names = set(zin.namelist())
                for item in zin.infolist():
                    fn = item.filename
                    data = zin.read(fn)
                    if fn == "[Content_Types].xml" and b"/xl/metadata.xml" not in data:
                        data = data.replace(
                            b"</Types>",
                            b'<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>',
                        )
                    elif fn == "xl/_rels/workbook.xml.rels" and b"sheetMetadata" not in data:
                        existing = [int(x) for x in re.findall(rb'Id="rId(\d+)"', data)]
                        next_id = max(existing, default=0) + 1
                        data = data.replace(
                            b"</Relationships>",
                            b'<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>' % next_id,
                        )
                    elif fn.startswith("xl/worksheets/") and b'<f t="array"' in data:
                        data = re.sub(
                            rb'(<c\b[^>]*?)(><f t="array")',
                            lambda m: m.group(1)
                            + (b"" if b'cm="' in m.group(1) else b' cm="1"')
                            + m.group(2),
                            data,
                        )
                    zout.writestr(item, data)
                if "xl/metadata.xml" not in names:
                    zout.writestr("xl/metadata.xml", metadata_xml)
        shutil.move(tmp_path, saved_path)
    finally:
        if os.path.isfile(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def _inject_external_links(original_path, saved_path, cache=None):
    """
    openpyxl save() corrupts external link XML files and their relationship
    entries. This function wholesale-replaces all external-link-related content
    in the saved xlsx zip with pristine copies.

    If cache (from _read_external_link_cache) is provided, use cached data
    instead of re-reading original_path (which may itself be corrupted).
    """
    if not os.path.isfile(saved_path):
        return

    if cache is not None:
        extlink_files = cache.get("files", {})
        extlink_rel_blocks = cache.get("rel_blocks", [])
        extlink_ct_blocks = cache.get("ct_blocks", [])
        if not extlink_files:
            return
    else:
        cache = _read_external_link_cache(original_path)
        if cache is None:
            return
        extlink_files = cache["files"]
        extlink_rel_blocks = cache["rel_blocks"]
        extlink_ct_blocks = cache["ct_blocks"]

    # Rebuild saved zip: drop corrupted external link entries, inject originals
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(tmp_fd)

    try:
        with zipfile.ZipFile(saved_path, 'r') as saved_zip:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                for item in saved_zip.infolist():
                    fn = item.filename
                    data = saved_zip.read(fn)

                    # Skip corrupted external link files (replace with originals later)
                    if fn.startswith("xl/externalLinks/") or fn.startswith("externalLinks/"):
                        continue

                    # Strip corrupted external-link entries from rels, inject originals
                    if fn == "xl/_rels/workbook.xml.rels" and extlink_rel_blocks:
                        data = _fix_rels_xml(data, extlink_rel_blocks)
                    elif fn == "[Content_Types].xml" and extlink_ct_blocks:
                        data = _fix_ct_xml(data, extlink_ct_blocks)

                    new_zip.writestr(item, data)

                # Write pristine external link files from original
                for name, content in extlink_files.items():
                    new_zip.writestr(name, content)

        shutil.move(tmp_path, saved_path)
    finally:
        if os.path.isfile(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def _fix_rels_xml(saved_xml_bytes, orig_extlink_blocks):
    """Remove corrupted external-link Relationship entries from saved rels
    XML, then inject pristine originals before </Relationships>."""
    import re
    xml_str = saved_xml_bytes.decode("utf-8")

    # Remove all external-link Relationship elements (case-insensitive)
    xml_str = re.sub(
        r'\s*<Relationship\s[^>]*Type="[^"]*externalLink[^"]*"[^>]*/?>',
        '', xml_str, flags=re.IGNORECASE
    )

    # Inject original external-link blocks before closing tag
    inject = "\n".join(orig_extlink_blocks)
    xml_str = xml_str.replace("</Relationships>", inject + "\n</Relationships>")

    return xml_str.encode("utf-8")


def _fix_ct_xml(saved_xml_bytes, orig_extlink_blocks):
    """Remove corrupted external-link Override entries from saved
    [Content_Types].xml, then inject pristine originals before </Types>."""
    import re
    xml_str = saved_xml_bytes.decode("utf-8")

    # Remove all external-link Override elements
    xml_str = re.sub(
        r'\s*<Override\s[^>]*PartName="[^"]*externalLink[^"]*"[^>]*/?>',
        '', xml_str, flags=re.IGNORECASE
    )

    # Inject original external-link blocks before closing tag
    inject = "\n".join(orig_extlink_blocks)
    xml_str = xml_str.replace("</Types>", inject + "\n</Types>")

    return xml_str.encode("utf-8")


# ---------------------------------------------------------------------------
# Merge resolver
# ---------------------------------------------------------------------------


def _resolve_rowname(a_val, b_val):
    """Get effective RowName from col A and col B values.
    If col A is formula, use col B value as the real RowName."""
    if a_val is not None:
        rn = str(a_val).strip()
        if rn and not rn.startswith("="):
            return rn
    if b_val is not None:
        rn = str(b_val).strip()
        if rn and not rn.startswith("="):
            return rn
    # Fallback: formula text from col A
    if a_val is not None:
        rn = str(a_val).strip()
        if rn.startswith("="):
            return rn
    return ""


def _read_preview_data_row(ws, row_idx, max_data_col):
    """Read C+ data cells from a preview row into canonical data columns."""
    row_data = {}
    for col_idx in range(1, max_data_col + 1):
        cell_val = ws.cell(row=row_idx, column=col_idx + 2).value
        if isinstance(cell_val, str) and cell_val.startswith("'="):
            cell_val = cell_val[1:]
        if isinstance(cell_val, str) and chr(0x2192) in cell_val and cell_val.startswith("="):
            parts = cell_val.split(f" {chr(0x2192)} ", 1)
            cell_val = parts[0] if parts[0].startswith("=") else cell_val
        row_data[col_idx] = cell_val
    return row_data


def _is_conflict_preview_cell(cell):
    """Whether a preview final cell was marked as a conflict cell."""
    color = getattr(getattr(cell, "fill", None), "fgColor", None)
    rgb = getattr(color, "rgb", None)
    return isinstance(rgb, str) and rgb.upper().endswith("FFCCCC")


def _adjust_array_formula_cell(cell_val, col_idx, new_row, row_map):
    """Shift an ArrayFormula's internal references and rebind its ref.

    Rows shift when the merge inserts rows above an existing row; without
    rebinding, the stale ref points at the inserted row and the formula
    evaluates the wrong data (or Excel rejects the workbook).
    """
    text = cell_val.text
    if isinstance(text, str) and text.startswith("="):
        text = adjust_formula(text, row_map)
    return ArrayFormula(f"{get_column_letter(col_idx)}{new_row}", text)


def _theirs_cell_is_array_formula(t_wb, sn, rowname, col_idx, theirs_colA_order):
    """Whether the source (theirs) cell for a new row is an array formula."""
    if t_wb is None or sn not in t_wb.sheetnames:
        return False
    t_pos = theirs_colA_order.get(rowname[0]) if rowname else None
    if t_pos is None:
        return False
    try:
        return isinstance(
            t_wb[sn].cell(row=t_pos + DATA_START_ROW, column=col_idx).value,
            ArrayFormula,
        )
    except Exception:
        return False


class MergeResolver:
    """
    Read decisions from merge preview file and apply them.
    - Extracts [最终] rows
    - Adjusts formula references
    - Writes resolved file
    """

    LABEL_MINE = "[我的]"
    LABEL_THEIRS = "[仓库]"
    LABEL_FINAL = "[最终]"

    def __init__(self, preview_path, output_path, mine_path=None, theirs_path=None, base_path=None):
        self.preview_path = preview_path
        self.output_path = output_path
        self.mine_path = mine_path
        self.theirs_path = theirs_path
        self.base_path = base_path
        # Workbooks opened during resolve(); closed in resolve() finally.
        self._open_wbs = []

    def resolve(self):
        """Read decisions from the merge preview and apply them.

        Ensures every workbook opened along the way is closed, even on
        error (leaked handles lock .mine/.r* files on Windows).
        """
        try:
            return self._resolve_apply()
        finally:
            for wb in self._open_wbs:
                try:
                    if wb is not None:
                        wb.close()
                except Exception:
                    pass
            self._open_wbs = []

    def _resolve_apply(self):
        """
        Merge [最终] decisions back into the complete mine file.
        Preserves all non-conflict rows, sheets, and formulas.
        Returns the output path.
        """
        # Cache external link data from pristine mine BEFORE openpyxl touches it
        extlink_cache = None
        if self.mine_path and os.path.isfile(self.mine_path):
            extlink_cache = _read_external_link_cache(self.mine_path)

        try:
            preview_wb = openpyxl.load_workbook(self.preview_path, data_only=False)
        except Exception as e:
            raise IOError(f"Cannot open preview file: {e}")
        self._open_wbs.append(preview_wb)

        # Load mine file as base (all sheets, all rows preserved)
        if self.mine_path:
            try:
                out_wb = _open_wb_any_ext(self.mine_path, data_only=False)
            except Exception as e:
                raise IOError(f"Cannot open mine file: {e}")
        else:
            # Fallback: try to infer mine path from preview path
            import re
            mine_candidate = re.sub(r'_merge\.xlsx$', '.xlsx', self.preview_path)
            if os.path.isfile(mine_candidate):
                try:
                    out_wb = _open_wb_any_ext(mine_candidate, data_only=False)
                except Exception:
                    out_wb = openpyxl.Workbook()
            else:
                out_wb = openpyxl.Workbook()
        self._open_wbs.append(out_wb)

        # Parse preview: for each sheet, read [最终] rows and their B-column decisions.
        # Col layout: A=[标记], B=[修改类型], C=RowName, D+=data
        # Source row ([我的]/[仓库]) key from col C (RowName) and col D (col B value).
        output_external_refs = _external_link_index_map(self.mine_path) if self.mine_path else {}
        output_external_name_map = {name: idx for idx, name in output_external_refs.items()}

        final_data = {}
        deleted_keys = {}  # {sheet: set of (colA, colB) keys to delete}
        preview_headers = {}
        for sn in preview_wb.sheetnames:
            ws = preview_wb[sn]
            max_data_col = ws.max_column - 2
            preview_headers[sn] = {
                "row1": [ws.cell(row=1, column=col_idx + 2).value for col_idx in range(1, max_data_col + 1)],
                "row2": [ws.cell(row=2, column=col_idx + 2).value for col_idx in range(1, max_data_col + 1)],
                "row3": [ws.cell(row=3, column=col_idx + 2).value for col_idx in range(1, max_data_col + 1)],
                "max_col": max_data_col,
            }
            sheet_finals = {}
            sheet_deleted = set()
            group_rownames = []
            group_sources = {}

            for row_idx in range(DATA_START_ROW, ws.max_row + 1):
                label = str(ws.cell(row=row_idx, column=1).value or "").strip()

                # Blank separator row → new conflict group
                if not label:
                    group_rownames.clear()
                    group_sources.clear()
                    continue

                if label in (self.LABEL_MINE, self.LABEL_THEIRS):
                    # Source row key: C=RowName, D=colB (was B,C before B column added)
                    col_a = _strip_col(ws.cell(row=row_idx, column=3).value)
                    col_b = _strip_col(ws.cell(row=row_idx, column=4).value)
                    if col_a:
                        group_rownames.append((col_a, col_b))
                    group_sources[label] = _read_preview_data_row(ws, row_idx, max_data_col)
                    continue

                if label != self.LABEL_FINAL:
                    continue

                # Read decision from B column
                decision = str(ws.cell(row=row_idx, column=2).value or "").strip()
                decision_norm = decision.lower()

                # Get identity key from the first source row in this group
                key = group_rownames[0] if group_rownames else None

                if decision == DECISION_DELETE:
                    if key:
                        sheet_deleted.add(key)
                    group_rownames.clear()
                    group_sources.clear()
                    continue

                if key is None:
                    group_rownames.clear()
                    group_sources.clear()
                    continue

                # Read data cells from the [最终] row.
                # Preserve formulas so auto-increment (=B49+1) and
                # cross-cell references (=I&J) stay intact.
                row_data = _read_preview_data_row(ws, row_idx, max_data_col)

                if decision_norm in (
                    DECISION_CONFLICT_MINE.lower(),
                    DECISION_CONFLICT_THEIRS.lower(),
                    DECISION_CONFLICT_EMPTY.lower(),
                ):
                    conflict_cols = [
                        col_idx for col_idx in range(1, max_data_col + 1)
                        if _is_conflict_preview_cell(ws.cell(row=row_idx, column=col_idx + 2))
                    ]
                    if decision_norm == DECISION_CONFLICT_EMPTY.lower():
                        for col_idx in conflict_cols:
                            row_data[col_idx] = ""
                    else:
                        source_label = self.LABEL_MINE if decision_norm == DECISION_CONFLICT_MINE.lower() else self.LABEL_THEIRS
                        source_data = group_sources.get(source_label, {})
                        for col_idx in conflict_cols:
                            row_data[col_idx] = source_data.get(col_idx, "")

                if key in sheet_finals:
                    print(f"[WARN] Duplicate RowName key {key} in preview sheet '{sn}'; later [最终] row overwrites earlier one.")
                sheet_finals[key] = row_data
                group_rownames.clear()
                group_sources.clear()

            final_data[sn] = sheet_finals
            if sheet_deleted:
                deleted_keys[sn] = sheet_deleted

        preview_wb.close()

        # Load theirs file for neighbor lookup
        # Build theirs_order using SAME (colA, colB) key scheme as rowname_to_row
        theirs_order = {}
        t_wb = None
        if self.theirs_path and os.path.isfile(self.theirs_path):
            try:
                t_wb = _open_wb_any_ext(self.theirs_path, data_only=False)
            except Exception:
                t_wb = None
        self._open_wbs.append(t_wb)

        # Load mine + theirs with data_only=True for VALUE-based key matching.
        # Formula cells need Excel-cached values. Without cached values,
        # fallback to formula-string keys (but matching may be less accurate).
        mine_val_wb = None
        t_val_wb = None
        try:
            if self.mine_path and os.path.isfile(self.mine_path):
                mine_val_wb = _open_wb_any_ext(self.mine_path, data_only=True, read_only=True)
        except Exception:
            mine_val_wb = None
        self._open_wbs.append(mine_val_wb)
        try:
            if self.theirs_path and os.path.isfile(self.theirs_path):
                t_val_wb = _open_wb_any_ext(self.theirs_path, data_only=True, read_only=True)
        except Exception:
            t_val_wb = None
        self._open_wbs.append(t_val_wb)

        # Apply [最终] values to the mine workbook
        # Iterate the union of out_wb and preview sheets: sheets added by
        # theirs (missing in mine) must not be silently dropped.
        all_out_sheets = list(out_wb.sheetnames) + [
            s for s in final_data if s not in out_wb.sheetnames
        ]
        # COM 直写计划：{sn: {created, headers, inserts, updates, deletes,
        # delete_from}}；仅 COM 路径使用，openpyxl 降级路径不收集
        com_plan = {}
        for sn in all_out_sheets:
            if not sn.startswith("#"):
                continue
            if sn not in final_data:
                continue
            if sn not in out_wb.sheetnames and not final_data[sn]:
                # Sheet missing in mine with no [最终] rows to write (e.g.
                # mine deleted the whole sheet) - do not resurrect it.
                continue

            sheet_created = False
            if sn not in out_wb.sheetnames:
                # Theirs-only sheet: create it; headers are written from the
                # preview (canonical theirs layout) by the block below.
                ws = out_wb.create_sheet(title=sn)
                sheet_created = True
            else:
                ws = out_wb[sn]
            plan_headers = {}
            preview_hdr = preview_headers.get(sn)
            if preview_hdr is not None:
                max_preview_data_col = max(0, preview_hdr.get("max_col", 0))
                for col_idx in range(1, max_preview_data_col + 1):
                    if sheet_created or col_idx > ws.max_column:
                        for header_row in (1, 2, 3):
                            values = preview_hdr.get(f"row{header_row}", [])
                            hval = values[col_idx - 1] if col_idx - 1 < len(values) else None
                            ws.cell(
                                row=header_row,
                                column=col_idx,
                                value=hval,
                            )
                            if hval is not None:
                                plan_headers[(header_row, col_idx)] = hval
            com_plan[sn] = {
                "created": sheet_created,
                "headers": plan_headers,
                "inserts": [],
                "updates": [],
                "deletes": [],
                "delete_from": None,
            }
            sheet_finals = final_data[sn]

            # Build mine key maps from COMPUTED VALUES (data_only=True).
            # These match _strip_col which also extracts computed values.
            rowname_to_row = {}
            colA_to_row = {}
            mine_colB = {}
            row_val_key = {}  # old_row → (colA_value, colB_value)
            if mine_val_wb and sn in mine_val_wb.sheetnames:
                mv_ws = mine_val_wb[sn]
                for row_idx, row in enumerate(
                    mv_ws.iter_rows(min_row=DATA_START_ROW, values_only=True),
                    start=DATA_START_ROW
                ):
                    if row and len(row) >= 1:
                        a_val = str(row[0] or "").strip()
                        b_val = str(row[1] or "").strip() if len(row) >= 2 else ""
                        if a_val:
                            rowname_to_row[(a_val, b_val)] = row_idx
                            colA_to_row[a_val] = row_idx
                            mine_colB[a_val] = b_val
                            row_val_key[row_idx] = (a_val, b_val)
            # 键来自 data_only 计算值时 RowName 值即行身份；公式字符串回退键
            # 可能出现相同文本对应不同行，不能仅凭 colA 判定身份。
            keys_from_values = bool(rowname_to_row)
            if not rowname_to_row:
                # Fallback: use formula-string keys from out_wb,
                # normalized via _strip_col for consistency.
                for row_idx in range(DATA_START_ROW, ws.max_row + 1):
                    a_val = _strip_col(ws.cell(row=row_idx, column=1).value)
                    b_val = _strip_col(ws.cell(row=row_idx, column=2).value)
                    if a_val:
                        rowname_to_row[(a_val, b_val)] = row_idx
                        colA_to_row[a_val] = row_idx
                        mine_colB[a_val] = b_val
                        row_val_key[row_idx] = (a_val, b_val)

            # Build theirs position map from COMPUTED VALUES.
            theirs_order = {}
            theirs_colA_order = {}
            theirs_pos_to_colA = {}
            if t_val_wb and sn in t_val_wb.sheetnames:
                tv_ws = t_val_wb[sn]
                for tr, row in enumerate(
                    tv_ws.iter_rows(min_row=DATA_START_ROW, values_only=True),
                    start=DATA_START_ROW
                ):
                    if row and len(row) >= 1:
                        ta = str(row[0] or "").strip()
                        tb = str(row[1] or "").strip() if len(row) >= 2 else ""
                        if ta:
                            pos = tr - DATA_START_ROW
                            theirs_order[(ta, tb)] = pos
                            theirs_colA_order[ta] = pos
                            theirs_pos_to_colA[pos] = ta
            if not theirs_order and t_wb and sn in t_wb.sheetnames:
                t_ws = t_wb[sn]
                for tr in range(DATA_START_ROW, t_ws.max_row + 1):
                    ta = _strip_col(t_ws.cell(row=tr, column=1).value)
                    tb = _strip_col(t_ws.cell(row=tr, column=2).value)
                    if ta:
                        pos = tr - DATA_START_ROW
                        theirs_order[(ta, tb)] = pos
                        theirs_colA_order[ta] = pos
                        theirs_pos_to_colA[pos] = ta

            # Collect cell values for ALL existing rows (for in-place rewrite)
            # Use iter_rows for batch read — ~10x faster than ws.cell()
            existing_rows = {}
            for row_idx, row_cells in enumerate(
                ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row,
                             max_col=ws.max_column),
                start=DATA_START_ROW
            ):
                existing_rows[row_idx] = {c + 1: cell.value for c, cell in enumerate(row_cells)}

            # Track deleted rows for Phase 4 filtering
            sheet_deleted = deleted_keys.get(sn, set())
            deleted_row_nums = set()  # mine row indices to skip in Phase 4
            for dk in sheet_deleted:
                if dk in rowname_to_row:
                    deleted_row_nums.add(rowname_to_row[dk])
                elif dk[0] in colA_to_row:
                    deleted_row_nums.add(colA_to_row[dk[0]])
            com_plan[sn]["deletes"] = sorted(deleted_row_nums)

            # Phase 1: Determine insert positions for new rows
            # A row in sheet_finals that doesn't exist in mine → new insert
            new_rows = []  # list of (insert_row_index, rowname, row_data)

            for rowname, row_data in sheet_finals.items():
                col_a, col_b = rowname

                # NEW row: exact (colA, colB) key not in mine → insert
                # Use full key to avoid false matches when formula-based colA
                # strings collide (same formula text, different computed values).
                if (col_a, col_b) in rowname_to_row:
                    continue  # exact match → existing row
                if col_a in colA_to_row and col_b == mine_colB.get(col_a, ""):
                    continue  # same colA + same colB → existing row (colB consistent)
                if keys_from_values and col_a in colA_to_row:
                    # RowName 值即行身份（重复 RowName 属于导表错误）。colA 已
                    # 存在于 mine 的 [最终] 行是对该行的修改（如 theirs 改了 B 列
                    # 注释导致 (colA, colB) 键匹配不上），绝非新增行——插入会造
                    # 成重复 RowName。Phase 4 会按 colA 回退匹配并把 [最终] 值
                    # 写到既有行上。
                    continue

                # Determine insertion point using colA-only neighbor matching
                insert_at = None
                if rowname[0] in theirs_colA_order:
                    target_pos = theirs_colA_order[rowname[0]]
                    max_pos = max(theirs_pos_to_colA.keys()) if theirs_pos_to_colA else 0
                    for offset in range(1, max_pos + 1):
                        prev_pos = target_pos - offset
                        if prev_pos >= 0 and prev_pos in theirs_pos_to_colA:
                            prev_ca = theirs_pos_to_colA[prev_pos]
                            if prev_ca in colA_to_row:
                                insert_at = colA_to_row[prev_ca] + 1
                                break
                        next_pos = target_pos + offset
                        if next_pos <= max_pos and next_pos in theirs_pos_to_colA:
                            next_ca = theirs_pos_to_colA[next_pos]
                            if next_ca in colA_to_row:
                                insert_at = colA_to_row[next_ca]
                                break

                if insert_at is None or insert_at < DATA_START_ROW:
                    insert_at = max(existing_rows.keys(), default=DATA_START_ROW - 1) + 1

                # Avoid collision: offset if position already taken by another new row
                while any(p[0] == insert_at for p in new_rows):
                    insert_at += 1

                new_rows.append((insert_at, rowname, row_data))

            # Phase 2: Build ordered row sequence (filter deleted, merge new + existing)
            # Each entry: (is_new, output_row, old_row|insert_at, rowname|None, row_data|None)
            new_rows.sort(key=lambda x: (x[0], x[1]))

            ordered = []
            new_idx = 0
            out_pos = DATA_START_ROW
            for old_row in sorted(existing_rows.keys()):
                if old_row in deleted_row_nums:
                    continue
                effective_pos = old_row
                while new_idx < len(new_rows) and new_rows[new_idx][0] <= effective_pos:
                    insert_at, rowname, row_data = new_rows[new_idx]
                    ordered.append((True, out_pos, insert_at, rowname, row_data))
                    out_pos += 1
                    effective_pos += 1
                    new_idx += 1
                ordered.append((False, out_pos, old_row, None, None))
                out_pos += 1
            while new_idx < len(new_rows):
                insert_at, rowname, row_data = new_rows[new_idx]
                ordered.append((True, out_pos, insert_at, rowname, row_data))
                out_pos += 1
                new_idx += 1

            # Phase 3: Build row_map from sequential output positions
            row_map = {}
            ext_row_map = {}
            for item in ordered:
                if item[0]:  # new row
                    _, out_row, _, rowname, _ = item
                    if rowname and rowname[0] in theirs_colA_order:
                        theirs_row = theirs_colA_order[rowname[0]] + DATA_START_ROW
                        ext_row_map[theirs_row] = out_row
                else:  # existing row
                    _, out_row, old_row, _, _ = item
                    row_map[old_row] = out_row
                    # Also map theirs row → output row for formula adjustment.
                    # THEIRS_MOD formulas reference theirs row numbers; this
                    # lets adjust_formula map them to the correct output row.
                    key = row_val_key.get(old_row, (None, None))
                    if key[0] and key[0] in theirs_colA_order:
                        theirs_row = theirs_colA_order[key[0]] + DATA_START_ROW
                        ext_row_map[theirs_row] = out_row

            # Phase 4: Write all rows sequentially (no gaps)
            for item in ordered:
                if item[0]:  # new row
                    _, write_row, _, rowname, row_data = item

                    # Clear all cells in this row first (openpyxl won't clear
                    # cell values by writing None — old data would leak through).
                    for col_idx in range(1, ws.max_column + 1):
                        ws.cell(row=write_row, column=col_idx).value = ""

                    # COM 直写计划：新行完整写入（最终行号），公式值由
                    # apply_write_plan_via_com 拒绝（design §4 / formula-write-reject）
                    plan_cells = {}
                    for col_idx, cell_val in row_data.items():
                        if isinstance(cell_val, str) and cell_val.startswith("="):
                            cell_val = _denormalize_external_formula_refs(cell_val, output_external_name_map)
                            cell_val = adjust_formula(cell_val, ext_row_map)
                            if _theirs_cell_is_array_formula(
                                t_wb, sn, rowname, col_idx, theirs_colA_order
                            ):
                                # New rows come from theirs; keep their array
                                # context, otherwise XLOOKUP degrades to
                                # implicit intersection (#VALUE!/first element).
                                cell_val = ArrayFormula(
                                    f"{get_column_letter(col_idx)}{write_row}",
                                    cell_val,
                                )
                        if cell_val is not None:
                            ws.cell(row=write_row, column=col_idx, value=cell_val)
                            plan_cells[col_idx] = cell_val
                    if plan_cells:
                        com_plan[sn]["inserts"].append((write_row, plan_cells))
                else:  # existing row
                    _, write_row, old_row, _, _ = item
                    new_row = write_row
                    old_vals = existing_rows[old_row]

                    # Use VALUE-based key from pre-built row_val_key map.
                    # Consistent with sheet_finals keys (from _strip_col VALUE).
                    key = row_val_key.get(old_row, ("", ""))
                    old_col_a = key[0]

                    has_final = key in sheet_finals
                    if not has_final and old_col_a:
                        for sk in sheet_finals:
                            if sk[0] == old_col_a:
                                has_final = True
                                key = sk
                                break
                    final_row_data = sheet_finals.get(key) if has_final else None

                    for col_idx in sorted(old_vals.keys()):
                        cell_val = old_vals[col_idx]
                        # COM 直写计划：记录最终要写入的单元格（公式值由
                        # apply 阶段拒绝；值未变 / mine 公式决策不写——行移动
                        # 的公式引用由 Excel 插入/删行时自动调整）
                        com_write = None

                        if col_idx == 1:
                            if isinstance(cell_val, str) and cell_val.startswith("="):
                                cell_val = adjust_formula(cell_val, row_map)
                            elif isinstance(cell_val, ArrayFormula):
                                cell_val = _adjust_array_formula_cell(cell_val, col_idx, new_row, row_map)
                            ws.cell(row=new_row, column=col_idx, value=cell_val if cell_val is not None else "")
                            continue

                        if final_row_data and col_idx in final_row_data:
                            final_val = final_row_data[col_idx]
                            if isinstance(final_val, str) and final_val.startswith("="):
                                final_val = _denormalize_external_formula_refs(final_val, output_external_name_map)
                                # Formula source: if same text as mine → mine formula
                                # (MINE_MOD) → use row_map. If different → theirs
                                # formula (THEIRS_MOD) → use ext_row_map.
                                old_cell = cell_val
                                old_text = (old_cell.text
                                            if isinstance(old_cell, ArrayFormula)
                                            else old_cell)
                                if isinstance(old_text, str) and old_text.startswith("=") \
                                        and old_text == final_val:
                                    # 决策 = mine 公式：COM 下无需重写，
                                    # Excel 插入/删行时自动调整引用
                                    cell_val = adjust_formula(final_val, row_map)
                                else:
                                    # 决策 = theirs 公式：记录待写，COM 路径
                                    # 落盘阶段拒绝（公式写入）
                                    cell_val = adjust_formula(final_val, ext_row_map)
                                    com_write = cell_val
                                if isinstance(old_cell, ArrayFormula):
                                    # Keep CSE array semantics: as a plain formula
                                    # string the cell loses array context and
                                    # evaluates with implicit intersection -> #VALUE!.
                                    cell_val = ArrayFormula(
                                        f"{get_column_letter(col_idx)}{new_row}",
                                        cell_val,
                                    )
                                    com_write = cell_val
                            else:
                                cell_val = final_val
                                # 静态值：与 mine 原值不同才写
                                if cell_val is not None and cell_val != old_vals.get(col_idx):
                                    com_write = cell_val
                        else:
                            if isinstance(cell_val, str) and cell_val.startswith("="):
                                cell_val = adjust_formula(cell_val, row_map)
                            elif isinstance(cell_val, ArrayFormula):
                                # ArrayFormula objects skip the str branch above;
                                # rows shifted by inserts must rebind refs too.
                                cell_val = _adjust_array_formula_cell(cell_val, col_idx, new_row, row_map)

                        ws.cell(row=new_row, column=col_idx, value=cell_val if cell_val is not None else "")
                        if com_write is not None:
                            com_plan[sn]["updates"].append((new_row, col_idx, com_write))

            # Clear trailing rows if output is shorter than original
            end_row = out_pos - 1
            # Delete trailing empty rows (not just clear — clear leaves
            # row count inflated, causing sheets to grow over time).
            if end_row < ws.max_row:
                ws.delete_rows(end_row + 1, ws.max_row - end_row)
                com_plan[sn]["delete_from"] = end_row + 1

        # Close read-only workbooks before saving output
        for wb in (preview_wb, t_wb, mine_val_wb, t_val_wb):
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass

        from excel_recalc import com_available, apply_write_plan_via_com
        if com_available():
            # COM 直写：必须先关闭 openpyxl 持有的 mine 句柄（文件锁）
            try:
                out_wb.close()
            except Exception:
                pass
            # COM 应用必须以 mine 原文件为基底；output 与 mine 不同（如
            # recover 用备份作 base）时先复制 mine 到 output
            if (self.mine_path and os.path.isfile(self.mine_path)
                    and os.path.abspath(str(self.mine_path))
                    != os.path.abspath(str(self.output_path))):
                shutil.copy2(str(self.mine_path), str(self.output_path))
            apply_write_plan_via_com(str(self.output_path), com_plan)
            return self.output_path

        # ---- 降级：openpyxl 全量写回 + zip 注入（无 Excel 环境） ----
        out_wb.save(self.output_path)

        # Close output workbook
        try:
            out_wb.close()
        except Exception:
            pass

        # Restore external links corrupted by openpyxl save
        if extlink_cache and extlink_cache.get("files"):
            _inject_external_links(self.mine_path, self.output_path, cache=extlink_cache)

        # Restore dynamic-array metadata dropped by openpyxl save
        _inject_dynamic_array_metadata(
            self.mine_path, self.output_path, fallback_path=self.theirs_path
        )

        print(
            "[resolve] 本机无 win32com/Excel，已降级 openpyxl 写回：公式缓存将被清空、"
            "外部链接需注入恢复、文件可能无法被 Excel 正常打开。"
            "建议在装有 Excel 的机器上重新执行合并。",
            file=sys.stderr,
        )

        return self.output_path


# ---------------------------------------------------------------------------
# Convenience functions
# ---------------------------------------------------------------------------

def generate_merge_preview(base_path, mine_path, theirs_path, output_path=None, conflict_default=DECISION_CONFLICT_THEIRS):
    """Generate merge preview for a conflicted file."""
    if output_path is None:
        # e.g., Effect.xlsx → Effect_merge.xlsx
        p = Path(mine_path)
        output_path = p.parent / f"{p.stem}_merge.xlsx"
    preview = MergePreview(base_path, mine_path, theirs_path, str(output_path), conflict_default=conflict_default)
    return preview.generate()


def _find_theirs_path(mine_path, preview_path):
    """Auto-detect theirs file path. Returns str or None."""
    import re

    candidates = []

    # 1) Infer from mine_path: replace _mine with _theirs
    if mine_path:
        d = os.path.dirname(str(mine_path))
        base_name = os.path.basename(str(mine_path))
        for pat in [('_mine.xlsx', '_theirs.xlsx'), ('_mine', '_theirs')]:
            candidate = os.path.join(d, base_name.replace(pat[0], pat[1]))
            if candidate != str(mine_path) and os.path.isfile(candidate):
                candidates.append(candidate)

    # 2) Try SVN conflict detection
    try:
        from svn_client import SvnClient
        svn = SvnClient()
        conflicts = svn.find_conflict_files()
        for base_rel, entry in conflicts.items():
            theirs = entry.get("theirs", "")
            if theirs and os.path.isfile(theirs):
                candidates.append(theirs)
    except Exception:
        pass

    # 3) Infer from preview_path
    candidate = re.sub(r'_merge\.xlsx$', '.xlsx', str(preview_path))
    if os.path.isfile(candidate) and (
        not mine_path
        or os.path.abspath(candidate) != os.path.abspath(str(mine_path))
    ):
        candidates.append(candidate)

    # Return first existing candidate
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def resolve_merge(preview_path, output_path=None, mine_path=None, theirs_path=None, base_path=None):
    """Resolve merge from preview file, using mine file as base for non-conflict data."""
    import re
    if output_path is None:
        p = Path(preview_path)
        stem = p.stem
        if stem.endswith("_merge"):
            original_stem = stem[:-6]
            output_path = p.parent / f"{original_stem}.xlsx"
        else:
            output_path = p.parent / f"{stem}_resolved.xlsx"
    if mine_path is None:
        mine_path = re.sub(r'_merge\.xlsx$', '.xlsx', str(preview_path))
    if theirs_path is None:
        theirs_path = _find_theirs_path(mine_path, preview_path)
    if base_path is None and mine_path:
        d = os.path.dirname(str(mine_path))
        b = os.path.join(d, os.path.basename(mine_path).replace('_mine', '_base'))
        if not os.path.isfile(b):
            b = os.path.join(d, os.path.basename(mine_path).replace('_mine.xlsx', '_base.xlsx'))
        if os.path.isfile(b):
            base_path = b
    resolver = MergeResolver(str(preview_path), str(output_path), mine_path, theirs_path, base_path)
    return resolver.resolve()
