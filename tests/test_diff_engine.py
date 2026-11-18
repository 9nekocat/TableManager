# -*- coding: utf-8 -*-
"""Unit tests for diff_engine value extraction, comparison, and display."""

import sys
from pathlib import Path
from collections import OrderedDict

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl

from diff_engine import (
    _build_diff_excel,
    _cell_compare_str,
    _cell_display,
    _compare_data,
    _extract_value,
    _normalize_dataset_formula_refs,
)

_ARROW = chr(0x2192)  # →


def test_extract_value_none():
    assert _extract_value(None) == ""


def test_extract_value_val_none():
    assert _extract_value((None, None)) == ""


def test_extract_value_pure_number():
    assert _extract_value((100, None)) == "100"


def test_extract_value_pure_string():
    assert _extract_value(("hello", None)) == "hello"


def test_extract_value_zero():
    assert _extract_value((0, None)) == "0"


def test_extract_value_format_display():
    val = f"=_xlfn.TEXTJOIN(...,[5]#Main!...,K7) {_ARROW} 0,3001,5.1"
    fml = "=_xlfn.TEXTJOIN(...)"
    assert _extract_value((val, fml)) == "0,3001,5.1"


def test_extract_value_format_display_no_space():
    val = f"=SUM(A1:A5){_ARROW}150"
    fml = "=SUM(A1:A5)"
    assert _extract_value((val, fml)) == "150"


def test_extract_value_format_display_multiple_arrows():
    """Only split on first arrow, value part may contain →."""
    val = f"=XLOOKUP(A1,B:B,C:C) {_ARROW} a {_ARROW} b"
    fml = "=XLOOKUP(A1,B:B,C:C)"
    assert _extract_value((val, fml)) == f"a {_ARROW} b"


def test_extract_value_no_arrow_in_val():
    """Regular cached value without format_display."""
    assert _extract_value(("0,3001,5.1", "=_xlfn.TEXTJOIN(...)")) == "0,3001,5.1"


# ── _cell_compare_str ────────────────────────────────────────────


def test_compare_same_pure_value():
    """Same clean values → no diff."""
    a = _cell_compare_str(("100", None))
    b = _cell_compare_str(("100", None))
    assert a == b


def test_compare_diff_pure_value():
    """Different values → diff."""
    a = _cell_compare_str(("100", None))
    b = _cell_compare_str(("200", None))
    assert a != b


def test_compare_formula_diff_value_same():
    """Formula text differs but computed value same → no diff."""
    old = _cell_compare_str(
        (f"=_xlfn.TEXTJOIN(...,[5]...,K7) {_ARROW} 0,3001,5.1",
         "=_xlfn.TEXTJOIN(...,[5]...,K7)")
    )
    new = _cell_compare_str(
        (f"=_xlfn.TEXTJOIN(...,[4]...,K7) {_ARROW} 0,3001,5.1",
         "=_xlfn.TEXTJOIN(...,[4]...,K7)")
    )
    assert old == new


def test_compare_one_format_display_one_clean():
    """One side has format_display, other has clean value — same value."""
    old = _cell_compare_str(
        (f"=SUM(A1:A5) {_ARROW} 150", "=SUM(A1:A5)")
    )
    new = _cell_compare_str(("150", None))
    assert old == new


def test_compare_both_none():
    assert _cell_compare_str(None) == _cell_compare_str((None, None))


# ── _cell_display ─────────────────────────────────────────────────


def test_display_clean_value():
    assert _cell_display(("hello", None)) == "hello"


def test_display_format_display_value():
    val = f"=_xlfn.TEXTJOIN(...) {_ARROW} Shield,2400"
    assert _cell_display((val, "=_xlfn.TEXTJOIN(...)")) == "Shield,2400"


def test_display_none():
    assert _cell_display(None) == ""


def test_display_val_none():
    assert _cell_display((None, "=SUM(A1:A5)")) == ""


def test_diff_excel_hides_unchanged_fields_for_modified_rows(tmp_path):
    headers = {
        "#Card": {
            "row1": tuple(f"字段{i}" for i in range(1, 9)),
            "row2": tuple(f"Field{i}" for i in range(1, 9)),
            "row3": tuple("string" for _ in range(8)),
            "max_col": 8,
        }
    }
    old_data = {
        "#Card": OrderedDict([
            ("Card_001", {
                1: ("Card_001", None),
                2: ("Name", None),
                3: ("same3", None),
                4: ("same4", None),
                5: ("old", None),
                6: ("same6", None),
                7: ("same7", None),
                8: ("same8", None),
            })
        ])
    }
    new_data = {
        "#Card": OrderedDict([
            ("Card_001", {
                1: ("Card_001", None),
                2: ("Name", None),
                3: ("same3", None),
                4: ("same4", None),
                5: ("new", None),
                6: ("same6", None),
                7: ("same7", None),
                8: ("same8", None),
            })
        ])
    }
    sheets_diff = _compare_data(old_data, new_data)
    output = tmp_path / "card_diff.xlsx"

    _build_diff_excel(sheets_diff, old_data, new_data, headers, output, "[旧版]", "[当前]")

    wb = openpyxl.load_workbook(output)
    ws = wb["#Card"]
    # A is the row label. Data columns 1, 2 and changed column 5 remain visible.
    assert ws.column_dimensions["B"].hidden is not True
    assert ws.column_dimensions["C"].hidden is not True
    assert ws.column_dimensions["F"].hidden is not True
    # Unchanged Card fields are still present for structure, but hidden by default.
    assert ws.column_dimensions["D"].hidden is True
    assert ws.column_dimensions["E"].hidden is True
    assert ws.column_dimensions["G"].hidden is True
    assert ws.column_dimensions["H"].hidden is True
    assert ws.column_dimensions["I"].hidden is True
    wb.close()


def test_card_diff_ignores_external_formula_slot_order_when_cache_missing(monkeypatch):
    data = {
        "#Card": OrderedDict([
            ("Card_001", {
                1: ("Card_001", None),
                2: (None, "=IF(A1=\"\",\"\",'[3]#AttributeTable'!A1)"),
            })
        ])
    }
    monkeypatch.setattr("diff_engine._external_link_index_map", lambda _path: {"3": "AttributeTable"})

    _normalize_dataset_formula_refs(data, "Card.xlsx")

    assert data["#Card"]["Card_001"][2][1] == "=IF(A1=\"\",\"\",'[AttributeTable]#AttributeTable'!A1)"

    old_data = {
        "#Card": OrderedDict([
            ("Card_001", {
                1: ("Card_001", None),
                2: (None, "=IF(A1=\"\",\"\",'[AttributeTable]#AttributeTable'!A1)"),
            })
        ])
    }
    assert _compare_data(old_data, data)["#Card"]["changes"] == []


def test_compare_data_empty_recalc_results_are_not_diffs():
    """After the pre-diff recalc, a None cache means an empty formula result.
    Empty-result formula vs empty cell (or vs another empty-result formula,
    e.g. shifted by an inserted row) must not be reported as a change."""
    old_data = {
        "#Main": OrderedDict([
            ("R1", {1: ("R1", None), 6: (None, None)}),
            ("R2", {1: ("R2", None), 6: (None, "=IF(E449=\"\",\"\",XLOOKUP(E449,A:A,B:B))")}),
            ("R3", {1: ("R3", None), 6: ("5", "=A1")}),
            ("R4", {1: ("R4", None), 6: (None, "=A1")}),
        ])
    }
    new_data = {
        "#Main": OrderedDict([
            # empty cell gained an empty-result formula → not a change
            ("R1", {1: ("R1", None), 6: (None, "=L2112&M2112")}),
            # same logical formula shifted by an inserted row → not a change
            ("R2", {1: ("R2", None), 6: (None, "=IF(E450=\"\",\"\",XLOOKUP(E450,A:A,B:B))")}),
            # cached value disappeared entirely → still reported
            ("R3", {1: ("R3", None), 6: (None, None)}),
            # formula changed and value became non-empty → reported
            ("R4", {1: ("R4", None), 6: ("7", "=A2")}),
        ])
    }
    changes = _compare_data(old_data, new_data)["#Main"]["changes"]
    by_row = {c["rowname"]: c for c in changes}
    assert "R1" not in by_row
    assert "R2" not in by_row
    assert by_row["R3"]["type"] == "modified"
    assert by_row["R4"]["type"] == "modified"
