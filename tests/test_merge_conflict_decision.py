import openpyxl

from merge_engine import generate_merge_preview, resolve_merge


def _write_book(path, value):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    headers = [
        ("RowName", "Name", "Value"),
        ("Row", "Name", "Value"),
        ("string", "string", "int"),
    ]
    for r_idx, row in enumerate(headers, start=1):
        for c_idx, cell_value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=cell_value)
    ws.cell(row=4, column=1, value="Hero_001")
    ws.cell(row=4, column=2, value="Alice")
    ws.cell(row=4, column=3, value=value)
    wb.save(path)
    wb.close()


def _preview_final_row(ws):
    for row in range(4, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == "[最终]":
            return row
    raise AssertionError("[最终] row not found")


def _resolve_value(preview, resolved, mine, theirs):
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))
    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        return wb["#Data"].cell(row=4, column=3).value
    finally:
        wb.close()


def test_conflict_final_defaults_to_theirs_with_dropdown(tmp_path):
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "preview_merge.xlsx"

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)

    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))

    wb = openpyxl.load_workbook(preview)
    ws = wb["#Data"]
    final_row = _preview_final_row(ws)
    formulas = [dv.formula1 for dv in ws.data_validations.dataValidation]

    assert ws.cell(row=final_row, column=2).value == "theirs"
    assert ws.cell(row=final_row, column=5).value == 300
    assert '"mine,theirs,空"' in formulas
    wb.close()


def test_conflict_dropdown_decisions_apply_on_resolve(tmp_path):
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "preview_merge.xlsx"

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))

    assert _resolve_value(preview, tmp_path / "resolved_theirs.xlsx", mine, theirs) == 300

    wb = openpyxl.load_workbook(preview)
    ws = wb["#Data"]
    final_row = _preview_final_row(ws)
    ws.cell(row=final_row, column=2, value="mine")
    wb.save(preview)
    wb.close()
    assert _resolve_value(preview, tmp_path / "resolved_mine.xlsx", mine, theirs) == 200

    wb = openpyxl.load_workbook(preview)
    ws = wb["#Data"]
    final_row = _preview_final_row(ws)
    ws.cell(row=final_row, column=2, value="空")
    wb.save(preview)
    wb.close()
    assert _resolve_value(preview, tmp_path / "resolved_empty.xlsx", mine, theirs) in ("", None)


def test_conflict_preview_default_strategy_can_be_selected(tmp_path):
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"

    _write_book(base, 100)
    _write_book(mine, 200)
    _write_book(theirs, 300)

    preview_mine = tmp_path / "preview_mine.xlsx"
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview_mine), conflict_default="mine")
    wb = openpyxl.load_workbook(preview_mine)
    ws = wb["#Data"]
    final_row = _preview_final_row(ws)
    assert ws.cell(row=final_row, column=2).value == "mine"
    assert ws.cell(row=final_row, column=5).value == 200
    wb.close()

    preview_empty = tmp_path / "preview_empty.xlsx"
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview_empty), conflict_default="empty")
    wb = openpyxl.load_workbook(preview_empty)
    ws = wb["#Data"]
    final_row = _preview_final_row(ws)
    assert ws.cell(row=final_row, column=2).value == "空"
    assert ws.cell(row=final_row, column=5).value in ("", None)
    wb.close()
