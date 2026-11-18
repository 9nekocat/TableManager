# -*- coding: utf-8 -*-
import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from diff_engine import diff_two_files
from merge_engine import generate_merge_preview, resolve_merge


DATA_START_ROW = 4


def _write_table(path, columns, rows, *, auto_filter=False, hidden_row_names=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"

    for col_idx, (cn, en, typ) in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=cn)
        ws.cell(row=2, column=col_idx, value=en)
        ws.cell(row=3, column=col_idx, value=typ)

    row_name_to_excel_row = {}
    for row_offset, row_data in enumerate(rows, start=DATA_START_ROW):
        row_name = row_data[columns[0][1]]
        row_name_to_excel_row[row_name] = row_offset
        for col_idx, (_, en, _) in enumerate(columns, start=1):
            ws.cell(row=row_offset, column=col_idx, value=row_data.get(en, ""))

    if auto_filter:
        ws.auto_filter.ref = f"A3:{openpyxl.utils.get_column_letter(len(columns))}{ws.max_row}"

    for row_name in hidden_row_names:
        ws.row_dimensions[row_name_to_excel_row[row_name]].hidden = True

    wb.save(path)
    wb.close()
    return path


def _header_col(ws, field_name):
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=2, column=col_idx).value == field_name:
            return col_idx
    raise AssertionError(f"missing field header: {field_name}")


def _row_by_key(ws, row_name):
    row_col = _header_col(ws, "Row")
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        if ws.cell(row=row_idx, column=row_col).value == row_name:
            return row_idx
    raise AssertionError(f"missing row: {row_name}")


def _cell(ws, row_name, field_name):
    return ws.cell(row=_row_by_key(ws, row_name), column=_header_col(ws, field_name)).value


def _rows_with_label(ws, label):
    return [
        row_idx
        for row_idx in range(DATA_START_ROW, ws.max_row + 1)
        if ws.cell(row=row_idx, column=1).value == label
    ]


@pytest.mark.parametrize("filtered_side", ["mine", "theirs"])
def test_merge_reads_filtered_and_unfiltered_workbooks(tmp_path, filtered_side):
    columns = [
        ("Row", "Row", "FString"),
        ("Name", "Name", "FString"),
        ("Power", "Power", "int32"),
        ("Tag", "Tag", "FString"),
    ]
    base_rows = [
        {"Row": "Item_001", "Name": "Alpha", "Power": 10, "Tag": "A"},
        {"Row": "Item_002", "Name": "Beta", "Power": 20, "Tag": "B"},
        {"Row": "Item_003", "Name": "Gamma", "Power": 30, "Tag": "C"},
    ]
    mine_rows = [dict(row) for row in base_rows]
    theirs_rows = [dict(row) for row in base_rows]
    theirs_rows[1]["Power"] = 22

    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"

    _write_table(base, columns, base_rows)
    _write_table(
        mine,
        columns,
        mine_rows,
        auto_filter=filtered_side == "mine",
        hidden_row_names={"Item_002"} if filtered_side == "mine" else set(),
    )
    _write_table(
        theirs,
        columns,
        theirs_rows,
        auto_filter=filtered_side == "theirs",
        hidden_row_names={"Item_002"} if filtered_side == "theirs" else set(),
    )

    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved)
    try:
        ws = wb["#Data"]
        assert _cell(ws, "Item_002", "Power") == 22
        assert _cell(ws, "Item_003", "Power") == 30
        if filtered_side == "mine":
            assert ws.auto_filter.ref == "A3:D6"
            assert ws.row_dimensions[5].hidden is True
        else:
            assert ws.auto_filter.ref is None
            assert ws.row_dimensions[5].hidden is False
    finally:
        wb.close()


def test_merge_aligns_moved_added_and_deleted_columns(tmp_path):
    base_columns = [
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("Name", "Name", "FString"),
        ("IconPath", "IconPath", "FString"),
        ("Power", "Power", "int32"),
        ("RemovedOnly", "RemovedOnly", "FString"),
        ("StableNote", "StableNote", "FString"),
    ]
    mine_columns = [
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("HelperAdded", "HelperAdded", "FString"),
        ("Power", "Power", "int32"),
        ("Name", "Name", "FString"),
        ("IconPath", "IconPath", "FString"),
        ("StableNote", "StableNote", "FString"),
    ]
    theirs_columns = base_columns[:-1] + [
        ("ServerOnly", "ServerOnly", "FString"),
        base_columns[-1],
    ]

    base_rows = [{
        "Row": "Card_001",
        "ID": 1,
        "Name": "Alpha",
        "IconPath": "icon/base.png",
        "Power": 10,
        "RemovedOnly": "legacy",
        "StableNote": "stable",
    }]
    mine_rows = [{
        "Row": "Card_001",
        "ID": 1,
        "Name": "Alpha",
        "HelperAdded": "helper_mine",
        "IconPath": "icon/base.png",
        "Power": 10,
        "StableNote": "stable",
    }]
    theirs_rows = [{
        "Row": "Card_001",
        "ID": 1,
        "Name": "Alpha",
        "IconPath": "icon/theirs.png",
        "Power": 11,
        "RemovedOnly": "legacy",
        "ServerOnly": "server_theirs",
        "StableNote": "stable",
    }]

    base = _write_table(tmp_path / "base.xlsx", base_columns, base_rows)
    mine = _write_table(tmp_path / "mine.xlsx", mine_columns, mine_rows)
    theirs = _write_table(tmp_path / "theirs.xlsx", theirs_columns, theirs_rows)
    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"

    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        ws = wb["#Data"]
        assert [_cell(ws, "Card_001", field) for field in ("HelperAdded", "Power", "IconPath", "ServerOnly")] == [
            "helper_mine",
            11,
            "icon/theirs.png",
            "server_theirs",
        ]
        assert _cell(ws, "Card_001", "RemovedOnly") in ("", None)
        assert _header_col(ws, "Power") < _header_col(ws, "Name") < _header_col(ws, "IconPath")
    finally:
        wb.close()


def _write_dup_header_table(path, theirs_desc):
    """5 columns where col 3/4 share an identical (cn, en, typ) header tuple,
    and col 5 holds a formula referencing the cell on its left."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    columns = [
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("标签", "None", ""),
        ("标签", "None", ""),
        ("Desc", "None", ""),
    ]
    for col_idx, (cn, en, typ) in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=cn)
        ws.cell(row=2, column=col_idx, value=en)
        ws.cell(row=3, column=col_idx, value=typ)
    values = ["Card_001", 1, "tagA", "tagB", theirs_desc]
    for col_idx, v in enumerate(values, start=1):
        ws.cell(row=DATA_START_ROW, column=col_idx, value=v)
    wb.save(path)
    wb.close()
    return path


def test_merge_duplicate_header_columns_do_not_shift_later_columns(tmp_path):
    """Regression: identical header tuples must not collapse into one canonical
    column — that shifted every later column left by one and turned relative
    formula references (e.g. =D4 in E4) into self-references."""
    base = _write_dup_header_table(tmp_path / "base.xlsx", "desc-base")
    mine = _write_dup_header_table(tmp_path / "mine.xlsx", "desc-base")
    theirs = _write_dup_header_table(tmp_path / "theirs.xlsx", "desc-theirs")
    # col 5 formula referencing its left neighbour, as in the real Card table
    for path in (base, mine, theirs):
        wb = openpyxl.load_workbook(path)
        ws = wb["#Data"]
        ws.cell(row=DATA_START_ROW, column=5, value="=D4")
        wb.save(path)
        wb.close()
    # theirs changes the dup-header col 4 value -> whole row counts as modified
    wb = openpyxl.load_workbook(theirs)
    ws = wb["#Data"]
    ws.cell(row=DATA_START_ROW, column=4, value="tagB2")
    wb.save(theirs)
    wb.close()

    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        ws = wb["#Data"]
        row = [ws.cell(row=DATA_START_ROW, column=c).value for c in range(1, 6)]
        assert row == ["Card_001", 1, "tagA", "tagB2", "=D4"]
        # No cell formula may reference its own address (circular reference).
        for excel_row in ws.iter_rows():
            for cell in excel_row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    assert cell.coordinate not in v.replace("$", "")
    finally:
        wb.close()


def test_merge_preserves_array_formula_on_modified_row(tmp_path):
    """Array formulas (openpyxl ArrayFormula) on merge-touched rows must survive
    as formulas instead of being flattened to their cached value."""
    from openpyxl.worksheet.formula import ArrayFormula

    def write(path, note):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "#Data"
        columns = [
            ("Row", "Row", "FString"),
            ("ID", "ID", "int32"),
            ("TagA", "TagA", "FString"),
            ("TagB", "TagB", "FString"),
            ("Tags", "Tags", "FString"),
            ("Note", "Note", "FString"),
        ]
        for col_idx, (cn, en, typ) in enumerate(columns, start=1):
            ws.cell(row=1, column=col_idx, value=cn)
            ws.cell(row=2, column=col_idx, value=en)
            ws.cell(row=3, column=col_idx, value=typ)
        ws.cell(row=DATA_START_ROW, column=1, value="Card_001")
        ws.cell(row=DATA_START_ROW, column=2, value=1)
        ws.cell(row=DATA_START_ROW, column=3, value="a")
        ws.cell(row=DATA_START_ROW, column=4, value="b")
        ws.cell(row=DATA_START_ROW, column=5).value = ArrayFormula(
            "E4", '=TEXTJOIN(",",TRUE,C4:D4)'
        )
        ws.cell(row=DATA_START_ROW, column=6, value=note)
        wb.save(path)
        wb.close()
        return path

    base = write(tmp_path / "base.xlsx", "base")
    mine = write(tmp_path / "mine.xlsx", "base")
    theirs = write(tmp_path / "theirs.xlsx", "theirs-changed")
    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"

    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        ws = wb["#Data"]
        v = ws.cell(row=DATA_START_ROW, column=5).value
        assert isinstance(v, ArrayFormula), f"array formula lost: {v!r}"
        assert "TEXTJOIN" in v.text, f"array formula flattened to {v!r}"
        assert ws.cell(row=DATA_START_ROW, column=6).value == "theirs-changed"
    finally:
        wb.close()


def _write_array_table(path, rows):
    """Write a #Data sheet with C/D tag columns and E=array TEXTJOIN formulas.

    rows: list of (rowname, tag_a, tag_b) — first entry lands on DATA_START_ROW.
    """
    from openpyxl.worksheet.formula import ArrayFormula

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    for col_idx, (cn, en, typ) in enumerate([
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("TagA", "TagA", "FString"),
        ("TagB", "TagB", "FString"),
        ("Tags", "Tags", "FString"),
    ], start=1):
        ws.cell(row=1, column=col_idx, value=cn)
        ws.cell(row=2, column=col_idx, value=en)
        ws.cell(row=3, column=col_idx, value=typ)
    for i, (rn, a, b) in enumerate(rows, start=DATA_START_ROW):
        ws.cell(row=i, column=1, value=rn)
        ws.cell(row=i, column=2, value=i - DATA_START_ROW + 1)
        ws.cell(row=i, column=3, value=a)
        ws.cell(row=i, column=4, value=b)
        ws.cell(row=i, column=5).value = ArrayFormula(
            f"E{i}", f'=TEXTJOIN(",",TRUE,C{i}:D{i})'
        )
    wb.save(path)
    wb.close()


def test_merge_preserves_array_formula_on_added_row(tmp_path):
    """theirs-added rows keep their array formula (Card.xlsx TEXTJOIN bug:
    a plain formula loses CSE context and evaluates with implicit intersection)."""
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    _write_array_table(str(base), [("Card_001", "a", "b")])
    _write_array_table(str(mine), [("Card_001", "a", "b")])
    _write_array_table(str(theirs), [("Card_001", "a", "b"), ("Card_002", "x", "y")])

    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        ws = wb["#Data"]
        v = ws.cell(row=DATA_START_ROW + 1, column=5).value
        assert isinstance(v, ArrayFormula), f"added-row array formula lost: {v!r}"
        assert "TEXTJOIN" in v.text and "C5:D5" in v.text, f"flattened to {v!r}"
    finally:
        wb.close()


def test_merge_rebinds_array_formula_refs_on_row_insert(tmp_path):
    """Rows shifted down by an inserted row must rebind array formula refs;
    a stale ref (e.g. Y4:AB4 on a cell now at row 5) evaluates the wrong row."""
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    _write_array_table(str(base), [("Card_001", "a", "b"), ("Card_002", "c", "d")])
    _write_array_table(str(mine), [("Card_001", "a", "b"), ("Card_002", "c", "d")])
    _write_array_table(str(theirs), [("Card_000", "z", "w"), ("Card_001", "a", "b"), ("Card_002", "c", "d")])

    preview = tmp_path / "preview.xlsx"
    resolved = tmp_path / "resolved.xlsx"
    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved), mine_path=str(mine), theirs_path=str(theirs))

    wb = openpyxl.load_workbook(resolved, data_only=False)
    try:
        ws = wb["#Data"]
        v_new = ws.cell(row=DATA_START_ROW, column=5).value
        assert isinstance(v_new, ArrayFormula), f"inserted-row array formula lost: {v_new!r}"
        assert "TEXTJOIN" in v_new.text and "C4:D4" in v_new.text, f"flattened to {v_new!r}"
        v_shifted = ws.cell(row=DATA_START_ROW + 1, column=5).value
        assert isinstance(v_shifted, ArrayFormula), f"shifted array formula lost: {v_shifted!r}"
        assert "TEXTJOIN" in v_shifted.text and "C5:D5" in v_shifted.text, f"refs not rebound: {v_shifted!r}"
    finally:
        wb.close()


_METADATA_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    b'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
    b'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
    b'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3">'
    b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    b'</ext></extLst></bk></futureMetadata>'
    b'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)


def _add_metadata_to_zip(path):
    """Simulate an Excel-saved workbook: inject xl/metadata.xml + cm="1" on
    array formula cells (openpyxl alone cannot write these)."""
    import re
    import shutil
    import tempfile

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>',
                )
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>',
                )
            elif item.filename.startswith("xl/worksheets/"):
                data = re.sub(
                    rb'(<c\b[^>]*?)(><f t="array")',
                    rb'\1 cm="1"\2',
                    data,
                )
            zout.writestr(item, data)
        zout.writestr("xl/metadata.xml", _METADATA_XML)
    shutil.move(tmp, path)


def test_inject_dynamic_array_metadata_restores_metadata_and_cm(tmp_path):
    """After an openpyxl save drops them, injection must restore
    xl/metadata.xml and the cm="1" attribute on array formula cells."""
    from openpyxl.worksheet.formula import ArrayFormula
    from merge_engine import _inject_dynamic_array_metadata

    source = tmp_path / "source.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    ws["A1"] = ArrayFormula("A1", '=TEXTJOIN(",",TRUE,B1:C1)')
    wb.save(source)
    wb.close()
    _add_metadata_to_zip(str(source))

    # openpyxl round-trip drops metadata.xml and cm attributes
    wb = openpyxl.load_workbook(source)
    saved = tmp_path / "saved.xlsx"
    wb.save(saved)
    wb.close()
    with zipfile.ZipFile(saved) as z:
        assert "xl/metadata.xml" not in z.namelist()

    _inject_dynamic_array_metadata(str(source), str(saved))

    with zipfile.ZipFile(saved) as z:
        assert z.read("xl/metadata.xml") == _METADATA_XML
        assert b"/xl/metadata.xml" in z.read("[Content_Types].xml")
        assert b"sheetMetadata" in z.read("xl/_rels/workbook.xml.rels")
        sheet = z.read("xl/worksheets/sheet1.xml")
        assert b'<c r="A1" cm="1"' in sheet or b'cm="1"' in sheet.split(b'<f t="array"')[0][-200:]


def test_diff_aligns_moved_added_and_deleted_columns(tmp_path, monkeypatch):
    columns_old = [
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("Name", "Name", "FString"),
        ("IconPath", "IconPath", "FString"),
        ("Power", "Power", "int32"),
        ("RemovedOnly", "RemovedOnly", "FString"),
        ("StableNote", "StableNote", "FString"),
    ]
    columns_new = [
        ("Row", "Row", "FString"),
        ("ID", "ID", "int32"),
        ("HelperAdded", "HelperAdded", "FString"),
        ("Power", "Power", "int32"),
        ("Name", "Name", "FString"),
        ("IconPath", "IconPath", "FString"),
        ("StableNote", "StableNote", "FString"),
    ]
    old_rows = [{
        "Row": "Card_001",
        "ID": 1,
        "Name": "Alpha",
        "IconPath": "icon/shared.png",
        "Power": 10,
        "RemovedOnly": "legacy",
        "StableNote": "stable",
    }]
    new_rows = [{
        "Row": "Card_001",
        "ID": 1,
        "Name": "Alpha",
        "HelperAdded": "helper_new",
        "IconPath": "icon/shared.png",
        "Power": 11,
        "StableNote": "stable",
    }]

    old_path = _write_table(tmp_path / "old.xlsx", columns_old, old_rows)
    new_path = _write_table(tmp_path / "new.xlsx", columns_new, new_rows)
    opened = []
    monkeypatch.setattr(os, "startfile", lambda path: opened.append(path), raising=False)
    monkeypatch.setattr("excel_recalc.recalc_batch", lambda filepaths: filepaths)

    output = diff_two_files(str(old_path), str(new_path), str(tmp_path), display_path=str(new_path))

    assert opened == [output]
    assert Path(output).exists()

    wb = openpyxl.load_workbook(output, data_only=False)
    try:
        ws = wb["#Data"]
        column_labels = [ws.cell(row=row_idx, column=1).value for row_idx in range(DATA_START_ROW, 10)]
        assert "[列新增]" in column_labels
        assert "[列删除]" in column_labels
        assert "[列移动]" in column_labels

        old_row = _rows_with_label(ws, "[旧版]")[0]
        new_row = _rows_with_label(ws, "[当前]")[0]

        assert ws.cell(row=2, column=_header_col(ws, "RemovedOnly")).value == "RemovedOnly"
        assert str(ws.cell(row=old_row, column=_header_col(ws, "Power")).value) == "10"
        assert str(ws.cell(row=new_row, column=_header_col(ws, "Power")).value) == "11"
        assert ws.cell(row=old_row, column=_header_col(ws, "HelperAdded")).value in ("", None)
        assert ws.cell(row=new_row, column=_header_col(ws, "HelperAdded")).value == "helper_new"
        assert ws.cell(row=old_row, column=_header_col(ws, "RemovedOnly")).value == "legacy"
        assert ws.cell(row=new_row, column=_header_col(ws, "RemovedOnly")).value in ("", None)
        assert ws.cell(row=old_row, column=_header_col(ws, "IconPath")).value == "icon/shared.png"
        assert ws.cell(row=new_row, column=_header_col(ws, "IconPath")).value == "icon/shared.png"
        assert ws.column_dimensions[openpyxl.utils.get_column_letter(_header_col(ws, "IconPath"))].hidden is True
    finally:
        wb.close()


def test_diff_column_addition_does_not_mark_every_row_modified(tmp_path, monkeypatch):
    old_columns = [
        ("Row", "Row", "FString"),
        ("Name", "Name", "FString"),
    ]
    new_columns = old_columns + [
        ("Tag", "Tag", "FString"),
    ]
    old_rows = [
        {"Row": "Item_001", "Name": "Alpha"},
        {"Row": "Item_002", "Name": "Beta"},
        {"Row": "Item_003", "Name": "Gamma"},
    ]
    new_rows = [
        {"Row": "Item_001", "Name": "Alpha", "Tag": "A"},
        {"Row": "Item_002", "Name": "Beta", "Tag": "B"},
        {"Row": "Item_003", "Name": "Gamma", "Tag": "C"},
    ]
    old_path = _write_table(tmp_path / "old.xlsx", old_columns, old_rows)
    new_path = _write_table(tmp_path / "new.xlsx", new_columns, new_rows)
    monkeypatch.setattr(os, "startfile", lambda path: None, raising=False)
    monkeypatch.setattr("excel_recalc.recalc_batch", lambda filepaths: filepaths)

    output = diff_two_files(str(old_path), str(new_path), str(tmp_path), display_path=str(new_path))

    wb = openpyxl.load_workbook(output, data_only=False)
    try:
        ws = wb["#Data"]
        labels = [ws.cell(row=row_idx, column=1).value for row_idx in range(DATA_START_ROW, ws.max_row + 1)]
        assert labels == ["[列新增]"]
        assert "新增列: Tag" in ws.cell(row=DATA_START_ROW, column=_header_col(ws, "Tag")).value
    finally:
        wb.close()
