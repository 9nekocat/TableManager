import os
from pathlib import Path

import openpyxl

from diff_engine import diff_two_files
from merge_engine import generate_merge_preview, resolve_merge


ROOT = Path(__file__).resolve().parents[1]
CASE_DIR = ROOT / "测试用例"
BASE = CASE_DIR / "TC16_列增删移动.xlsx.r1028"
MINE = CASE_DIR / "TC16_列增删移动.xlsx.mine"
THEIRS = CASE_DIR / "TC16_列增删移动.xlsx.r1029"


def _header_index(ws, field_name):
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=2, column=col).value == field_name:
            return col
    raise AssertionError(f"header not found: {field_name}")


def _header_index_by_cn(ws, cn_name):
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == cn_name:
            return col
    raise AssertionError(f"header not found: {cn_name}")


def _first_row_with_label(ws, label):
    for row in range(4, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == label:
            return row
    raise AssertionError(f"label not found: {label}")


def test_tc16_diff_tool_aligns_moved_columns_by_header(tmp_path, monkeypatch):
    opened = []
    monkeypatch.setattr(os, "startfile", lambda path: opened.append(path), raising=False)
    monkeypatch.setattr("excel_recalc.recalc_batch", lambda filepaths: filepaths)

    output = diff_two_files(str(BASE), str(MINE), str(tmp_path), display_path=str(MINE))

    assert Path(output).exists()
    assert opened == [output]

    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb["#TacticalEncounterData"]
    icon_col = _header_index(ws, "IconPath")
    style_col = _header_index(ws, "StyleID")
    helper_col = _header_index_by_cn(ws, "样式辅助列")
    old_row = _first_row_with_label(ws, "[旧版]")
    new_row = old_row + 1

    assert ws.cell(row=old_row, column=icon_col).value == "icon/base_a.png"
    assert ws.cell(row=new_row, column=icon_col).value == "icon/mine_a.png"
    assert ws.cell(row=old_row, column=style_col).value == "900"
    assert ws.cell(row=new_row, column=style_col).value == "900"
    assert ws.cell(row=new_row, column=helper_col).value == "helper_mine_a"
    wb.close()


def test_tc16_merge_aligns_theirs_values_under_mine_headers(tmp_path):
    preview = tmp_path / "TC16_列增删移动_merge.xlsx"
    resolved = tmp_path / "TC16_列增删移动_resolved.xlsx"

    generate_merge_preview(str(BASE), str(MINE), str(THEIRS), str(preview))

    wb = openpyxl.load_workbook(preview)
    ws = wb["#TacticalEncounterData"]
    icon_col = _header_index(ws, "IconPath")
    style_col = _header_index(ws, "StyleID")
    helper_col = _header_index_by_cn(ws, "样式辅助列")
    theirs_row = _first_row_with_label(ws, "[仓库]")
    final_row = _first_row_with_label(ws, "[最终]")

    assert ws.cell(row=theirs_row, column=icon_col).value == "icon/theirs_a.png"
    assert ws.cell(row=theirs_row, column=style_col).value == 900
    assert ws.cell(row=theirs_row, column=helper_col).value in (None, "")

    # Simulate the user choosing the repository IconPath while keeping mine's
    # new helper column and moved StyleID in their proper columns.
    ws.cell(row=final_row, column=icon_col, value="icon/theirs_a.png")
    wb.save(preview)
    wb.close()

    resolve_merge(str(preview), str(resolved), mine_path=str(MINE), theirs_path=str(THEIRS))

    out_wb = openpyxl.load_workbook(resolved, data_only=False)
    out_ws = out_wb["#TacticalEncounterData"]
    out_icon_col = _header_index(out_ws, "IconPath")
    out_style_col = _header_index(out_ws, "StyleID")
    out_helper_col = _header_index_by_cn(out_ws, "样式辅助列")

    assert out_ws.cell(row=4, column=out_icon_col).value == "icon/theirs_a.png"
    assert out_ws.cell(row=4, column=out_style_col).value == 900
    assert out_ws.cell(row=4, column=out_helper_col).value == "helper_mine_a"
    assert out_ws.max_column == 13
    out_wb.close()
