# -*- coding: utf-8 -*-
"""Unit tests for merge preview add/delete row display (Bug 1).

Validates that MergePreview correctly:
- Shows mine data for MINE_ADDED rows
- Shows theirs data for THEIRS_ADDED rows
- Shows base data for deleted rows with [已删除] label
- Defaults [最终] row to modifying side for non-conflict cases
"""

import sys
import os
import tempfile
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
from openpyxl.styles import PatternFill
from merge_engine import (
    ConflictDetector,
    MergePreview,
    generate_merge_preview,
    _normalize_external_formula_refs,
    rewrite_external_formula_refs_to_names,
)
from formula_utils import normalize_formula_for_compare

DATA_START_ROW = 4


def test_formula_compare_normalizes_relative_reference_row_shift():
    assert normalize_formula_for_compare("=B42+C$2+$D42+$E$5", 42, 1) == \
        "=RC[1]+R2C[2]+RC4+R5C5"
    assert normalize_formula_for_compare("=B43+C$2+$D43+$E$5", 43, 1) == \
        "=RC[1]+R2C[2]+RC4+R5C5"


def test_formula_compare_keeps_absolute_reference_changes_visible():
    assert normalize_formula_for_compare("=$B$42", 42, 1) != \
        normalize_formula_for_compare("=$B$43", 43, 1)


def test_conflict_detector_ignores_relative_formula_row_shift_without_cache(tmp_path):
    """A moved Card row's copied formula must not become a modification."""
    paths = [tmp_path / name for name in ("base.xlsx", "mine.xlsx", "theirs.xlsx")]

    def write(path, rows):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "#Card"
        for row_idx, names in enumerate((("RowName", "Name", "Derived"),) * 3, start=1):
            for col_idx, value in enumerate(names, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        for row_idx, row in enumerate(rows, start=DATA_START_ROW):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        wb.save(path)
        wb.close()

    # openpyxl writes formulas without Excel's cached results. The same card
    # moves down by one row in theirs, so Excel updates =B4 to =B5.
    write(paths[0], [("Card_001", "Sword", "=B4")])
    write(paths[1], [("Card_001", "Sword", "=B4")])
    write(paths[2], [("Card_000", "Shield", "=B4"), ("Card_001", "Sword", "=B5")])

    rows = ConflictDetector(*(str(path) for path in paths)).detect()["#Card"]["rows"]
    card_row = next(row for row in rows if row["rowname"] == "Card_001")
    assert card_row["status"] == "same"


def test_external_formula_slots_normalize_to_workbook_names():
    formula = (
        "=IF(AO2=\"\",\"\",_xlfn.XLOOKUP(AO2,'[5]#Main'!$C:$C,"
        "'[5]#Main'!$B:$B)+_xlfn.XLOOKUP(B4,[4]价值计算!$A:$A,[4]价值计算!$F:$F))"
    )

    normalized = _normalize_external_formula_refs(
        formula,
        {"4": "CardTable", "5": "Effect"},
    )

    assert "[4]" not in normalized
    assert "[5]" not in normalized
    assert "[CardTable]价值计算" in normalized
    assert "'[Effect]#Main'!" in normalized


def test_rewrite_external_formula_slots_in_workbook(tmp_path):
    path = tmp_path / "formula_slots.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    ws["A1"] = "Header"
    ws["A2"] = "Field"
    ws["A3"] = "Type"
    ws["A4"] = "=_xlfn.XLOOKUP(B4,[4]价值计算!$A:$A,[4]价值计算!$F:$F)"
    wb.save(path)
    wb.close()

    result = rewrite_external_formula_refs_to_names(path, {"4": "CardTable"})

    assert result["rewritten"] == 1
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        assert "[4]" not in wb["#Data"]["A4"].value
        assert "[CardTable]价值计算" in wb["#Data"]["A4"].value
    finally:
        wb.close()


def _create_minimal_wb(filename, rows_data):
    """Create a minimal workbook with one #SheetName sheet and given data rows.

    rows_data: list of (row_name_col_A, col_B, col_C, ...) for data rows (row 4+).
    Header rows 1-3 are auto-generated.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#TestSheet"

    # Row 1: Chinese headers
    headers_cn = ["RowName", "Name", "Value"]
    for i, h in enumerate(headers_cn, 1):
        ws.cell(row=1, column=i, value=h)

    # Row 2: English field names
    for i, h in enumerate(["RowName", "Name", "Value"], 1):
        ws.cell(row=2, column=i, value=h)

    # Row 3: Types
    for i, h in enumerate(["FString", "FString", "int32"], 1):
        ws.cell(row=3, column=i, value=h)

    # Data rows
    for ri, row_data in enumerate(rows_data, DATA_START_ROW):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val)

    wb.save(filename)
    wb.close()
    return filename


def _read_preview_rows(preview_path):
    """Parse merge preview xlsx back into structured data.

    Returns list of dicts: {row_num, label, change_type, col_a, col_b, col_c}
    """
    wb = openpyxl.load_workbook(preview_path, data_only=True)
    rows = []
    for sn in wb.sheetnames:
        if sn.startswith("#"):
            ws = wb[sn]
            for row_idx in range(DATA_START_ROW, ws.max_row + 1):
                label = str(ws.cell(row=row_idx, column=1).value or "").strip()
                if not label:
                    continue  # blank separator

                change_type = str(ws.cell(row=row_idx, column=2).value or "").strip()
                row_data = {}
                max_col = min(ws.max_column, 6)
                for col in range(3, max_col + 1):
                    row_data[f"col{col-2}"] = ws.cell(row=row_idx, column=col).value
                rows.append({
                    "row": row_idx,
                    "label": label,
                    "change_type": change_type,
                    "data": row_data,
                })
    wb.close()
    return rows


class TestAddDeletePreview:
    """Bug 1: Merge preview add/delete row display."""

    def setup_method(self):
        self.tmpdir = tempfile.mkdtemp()
        self.base_path = os.path.join(self.tmpdir, "base.xlsx")
        self.mine_path = os.path.join(self.tmpdir, "mine.xlsx")
        self.theirs_path = os.path.join(self.tmpdir, "theirs.xlsx")
        self.preview_path = os.path.join(self.tmpdir, "preview.xlsx")

    def teardown_method(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def _make_preview(self, base_rows, mine_rows, theirs_rows):
        _create_minimal_wb(self.base_path, base_rows)
        _create_minimal_wb(self.mine_path, mine_rows)
        _create_minimal_wb(self.theirs_path, theirs_rows)
        return generate_merge_preview(
            self.base_path, self.mine_path, self.theirs_path, self.preview_path
        )

    def test_preview_records_source_sheet_order_metadata(self):
        self._make_preview(
            [("Hero_001", "Alice", "100")],
            [("Hero_001", "Alice", "101")],
            [("Hero_001", "Alice", "100")],
        )

        wb = openpyxl.load_workbook(self.preview_path, data_only=True)
        try:
            assert "_merge_meta" in wb.sheetnames
            ws = wb["_merge_meta"]
            assert ws.sheet_state == "hidden"
            rows = list(ws.iter_rows(values_only=True))
            assert rows[0] == ("kind", "role", "path", "sheet_name", "sheet_index", "is_data_sheet")
            sheet_rows = [row for row in rows[1:] if row[0] == "sheet_order"]
            roles = {row[1] for row in sheet_rows}
            assert {"base", "mine", "theirs"}.issubset(roles)
            assert any(row[3] == "#TestSheet" and row[4] == 1 and row[5] is True for row in sheet_rows)
        finally:
            wb.close()

    # ---------------------------------------------------------------
    # MINE_ADDED: row not in base, exists in mine, not in theirs
    # ---------------------------------------------------------------
    def test_mine_added_shows_mine_data(self):
        """[我的] row for MINE_ADDED shows mine data, not blank."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),   # NEW: MINE_ADDED
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        # Find MINE_ADDED group: [我的] with [已添加]
        mine_rows_found = [r for r in rows if r["label"] == "[我的]" and r["change_type"] == "[已添加]"]
        assert len(mine_rows_found) == 1, f"Expected 1 MINE_ADDED [我的] row, got {len(mine_rows_found)}"
        mr = mine_rows_found[0]
        # col1 should be "Hero_002" (mine RowName), not empty
        assert mr["data"].get("col1") == "Hero_002", \
            f"MINE_ADDED RowName expected 'Hero_002', got '{mr['data'].get('col1')}'"
        # col2 should show value from mine ("Bob")
        assert mr["data"].get("col2") == "Bob", \
            f"MINE_ADDED Name expected 'Bob', got '{mr['data'].get('col2')}'"

    def test_mine_added_final_defaults_mine(self):
        """[最终] for MINE_ADDED pre-fills mine data and [修改] decision."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        final_rows = [r for r in rows if r["label"] == "[最终]"]
        # Should have at least 1 final row (one for Hero_002)
        fr_mine_added = [r for r in final_rows
                         if r["data"].get("col1") == "Hero_002"]
        assert len(fr_mine_added) == 1, f"Expected [最终] for Hero_002, got {len(fr_mine_added)}"
        fr = fr_mine_added[0]
        assert fr["change_type"] == "[修改]", \
            f"MINE_ADDED default should be [修改], got '{fr['change_type']}'"
        assert fr["data"].get("col2") == "Bob", \
            f"[最终] should pre-fill mine data 'Bob', got '{fr['data'].get('col2')}'"

    # ---------------------------------------------------------------
    # MINE_DELETED: row in base, NOT in mine, in theirs (mine deleted it)
    # ---------------------------------------------------------------
    def test_mine_deleted_shows_base_data(self):
        """[我的] for MINE_DELETED shows base data with [已删除] label."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),   # present in base
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            # Hero_002 removed by mine → MINE_DELETED
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),   # theirs still has it
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        # The MINE_DELETED row: [我的] with [已删除]
        mine_del = [r for r in rows if r["label"] == "[我的]" and r["change_type"] == "[已删除]"]
        assert len(mine_del) == 1, f"Expected 1 MINE_DELETED [我的] row, got {len(mine_del)}"
        md = mine_del[0]
        # Should show Hero_002 data from base
        assert md["data"].get("col1") == "Hero_002", \
            f"MINE_DELETED should show RowName 'Hero_002', got '{md['data'].get('col1')}'"

    def test_mine_deleted_final_shows_base_data(self):
        """[最终] for MINE_DELETED shows base data (mine side deleted, show what was)."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            # Hero_002 deleted by mine
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        final_rows = [r for r in rows if r["label"] == "[最终]"]
        fr_hero2 = [r for r in final_rows
                    if r["data"].get("col1") == "Hero_002"]
        assert len(fr_hero2) == 1, f"Expected [最终] for Hero_002, got {len(fr_hero2)}"
        fr = fr_hero2[0]
        assert fr["change_type"] == "[删除]", \
            f"MINE_DELETED default should be [删除], got '{fr['change_type']}'"
        # [最终] should show base data (fill_side=base), NOT None
        assert fr["data"].get("col2") == "Bob", \
            f"[最终] for MINE_DELETED should show base value 'Bob', got '{fr['data'].get('col2')}'"
        assert str(fr["data"].get("col3")) == "200", \
            f"[最终] for MINE_DELETED should show base value '200', got '{fr['data'].get('col3')}'"

    # ---------------------------------------------------------------
    # THEIRS_ADDED: row not in base, not in mine, exists in theirs
    # ---------------------------------------------------------------
    def test_theirs_added_shows_theirs_data(self):
        """[仓库] for THEIRS_ADDED shows theirs data."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_003", "Charlie", "300"),  # NEW: THEIRS_ADDED
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        theirs_added = [r for r in rows if r["label"] == "[仓库]" and r["change_type"] == "[已添加]"]
        assert len(theirs_added) == 1, f"Expected 1 THEIRS_ADDED [仓库] row, got {len(theirs_added)}"
        ta = theirs_added[0]
        assert ta["data"].get("col1") == "Hero_003", \
            f"THEIRS_ADDED RowName expected 'Hero_003', got '{ta['data'].get('col1')}'"
        assert ta["data"].get("col2") == "Charlie", \
            f"THEIRS_ADDED Name expected 'Charlie', got '{ta['data'].get('col2')}'"

    def test_theirs_added_final_defaults_theirs(self):
        """[最终] for THEIRS_ADDED pre-fills theirs data."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_003", "Charlie", "300"),
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        final_rows = [r for r in rows if r["label"] == "[最终]"]
        fr_hero3 = [r for r in final_rows
                    if r["data"].get("col1") == "Hero_003"]
        assert len(fr_hero3) == 1, f"Expected [最终] for Hero_003, got {len(fr_hero3)}"
        fr = fr_hero3[0]
        assert fr["change_type"] == "[修改]", \
            f"THEIRS_ADDED default should be [修改], got '{fr['change_type']}'"
        assert fr["data"].get("col2") == "Charlie", \
            f"[最终] should pre-fill theirs data 'Charlie', got '{fr['data'].get('col2')}'"

    # ---------------------------------------------------------------
    # THEIRS_DELETED: row in base and mine, NOT in theirs (theirs deleted)
    # ---------------------------------------------------------------
    def test_theirs_deleted_shows_base_data(self):
        """[仓库] for THEIRS_DELETED shows base data with [已删除]."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),   # mine still has it
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            # Hero_002 removed by theirs → THEIRS_DELETED
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        theirs_del = [r for r in rows if r["label"] == "[仓库]" and r["change_type"] == "[已删除]"]
        assert len(theirs_del) == 1, f"Expected 1 THEIRS_DELETED [仓库] row, got {len(theirs_del)}"
        td = theirs_del[0]
        assert td["data"].get("col1") == "Hero_002", \
            f"THEIRS_DELETED should show RowName 'Hero_002', got '{td['data'].get('col1')}'"

    def test_theirs_deleted_final_defaults_delete(self):
        """[最终] for THEIRS_DELETED defaults to [删除], pre-fills mine data."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),
            ("Hero_002", "Bob", "200"),   # mine still has it
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),
            # Hero_002 removed by theirs
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        final_rows = [r for r in rows if r["label"] == "[最终]"]
        fr_hero2 = [r for r in final_rows
                    if r["data"].get("col1") == "Hero_002"]
        assert len(fr_hero2) == 1
        fr = fr_hero2[0]
        assert fr["change_type"] == "[删除]", \
            f"THEIRS_DELETED default should be [删除], got '{fr['change_type']}'"
        # [最终] should show mine data (fill_side=mine, mine still has row)
        assert fr["data"].get("col2") == "Bob", \
            f"[最终] should pre-fill mine name 'Bob', got '{fr['data'].get('col2')}'"

    # ---------------------------------------------------------------
    # Non-conflict: THEIRS_MOD should default to theirs data in [最终]
    # ---------------------------------------------------------------
    def test_theirs_mod_final_prefills_theirs(self):
        """Non-conflict THEIRS_MOD [最终] defaults to theirs data."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),      # unchanged
        ]
        theirs_rows = [
            ("Hero_001", "Alice_Updated", "150"),  # THEIRS_MOD
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        final_rows = [r for r in rows if r["label"] == "[最终]"]
        fr_hero1 = [r for r in final_rows
                    if r["data"].get("col1") == "Hero_001"]
        assert len(fr_hero1) == 1, f"Expected [最终] for Hero_001, got {len(fr_hero1)}"
        fr = fr_hero1[0]
        # Should pre-fill theirs data (Alice_Updated, 150)
        assert fr["data"].get("col2") == "Alice_Updated", \
            f"[最终] should pre-fill theirs name 'Alice_Updated', got '{fr['data'].get('col2')}'"
        assert str(fr["data"].get("col3")) == "150", \
            f"[最终] should pre-fill theirs value '150', got '{fr['data'].get('col3')}'"

    # ---------------------------------------------------------------
    # SAME rows should NOT appear in preview
    # ---------------------------------------------------------------
    def test_same_rows_are_omitted(self):
        """SAME rows are not included in preview (only conflicts shown)."""
        base_rows = [
            ("Hero_001", "Alice", "100"),
        ]
        mine_rows = [
            ("Hero_001", "Alice", "100"),     # SAME
            ("Hero_002", "Bob", "200"),       # MINE_ADDED
        ]
        theirs_rows = [
            ("Hero_001", "Alice", "100"),     # SAME
        ]

        self._make_preview(base_rows, mine_rows, theirs_rows)
        rows = _read_preview_rows(self.preview_path)

        # Hero_001 (SAME) should not appear
        same_entries = [r for r in rows
                        if r["data"].get("col1") == "Hero_001"]
        assert len(same_entries) == 0, \
            f"SAME rows should be omitted from preview, found {len(same_entries)}"
