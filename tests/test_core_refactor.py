from __future__ import annotations

import json
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from core.diff import diff_files, format_text_diff, write_diff_excel
from core.formula import analyze_formula
from core.merge import (
    MergeConflictError,
    format_merge_text,
    merge_files,
    write_merge_excel,
)
from core.model import DATA_START_ROW, Header, Sheet, WorkbookDocument
from core.operations import apply_operations
from core_cli import main as cli_main


def _write_table(path, columns, rows, *, prefix_rows=0, blank_middle=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    for col_index, (cn, en, type_value) in enumerate(columns, start=1):
        ws.cell(row=1, column=col_index, value=cn)
        ws.cell(row=2, column=col_index, value=en)
        ws.cell(row=3, column=col_index, value=type_value)
    target_row = DATA_START_ROW + prefix_rows
    for row_index, row in enumerate(rows):
        if blank_middle and row_index == 1:
            target_row += 1
        for col_index, (_, en, _) in enumerate(columns, start=1):
            value = row.get(en)
            if value is not None:
                ws.cell(row=target_row, column=col_index, value=value)
        target_row += 1
    wb.save(path)
    wb.close()
    return path


def _table(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        ws = wb["#Data"]
        headers = [
            ws.cell(row=2, column=index).value
            for index in range(1, ws.max_column + 1)
        ]
        rows = [
            [ws.cell(row=row_index, column=col_index).value for col_index in range(1, ws.max_column + 1)]
            for row_index in range(DATA_START_ROW, ws.max_row + 1)
        ]
        return headers, rows
    finally:
        wb.close()


COLUMNS = [
    ("行名", "Row", "FString"),
    ("名称", "Name", "FString"),
    ("数值", "Power", "int32"),
]


def test_duplicate_row_keys_are_preserved_and_reported(tmp_path):
    path = _write_table(
        tmp_path / "duplicates.xlsx",
        COLUMNS,
        [
            {"Row": "ID_A", "Name": "First", "Power": 1},
            {"Row": "ID_A", "Name": "Second", "Power": 2},
            {"Row": "ID_B", "Name": "Third", "Power": 3},
        ],
    )

    document = WorkbookDocument.from_xlsx(path)
    sheet = document.sheet("#Data")
    assert len(sheet.find_rows("ID_A")) == 2
    assert [row.key for row in sheet.rows] == ["ID_A", "ID_A", "ID_B"]
    report = document.validate()
    assert any(issue.code == "duplicate_row_key" and issue.is_error for issue in report.issues)


def test_duplicate_headers_are_reported_without_collapsing_columns(tmp_path):
    path = _write_table(
        tmp_path / "headers.xlsx",
        [
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("名称2", "Name", "FString"),
        ],
        [{"Row": "A", "Name": "left"}],
    )
    document = WorkbookDocument.from_xlsx(path)
    sheet = document.sheet("#Data")
    assert len(sheet.headers) == 3
    report = document.validate()
    assert any(issue.code == "duplicate_english_header" for issue in report.issues)


def test_formula_risk_contract():
    assert not analyze_formula("=C4&D4", source_row=4, source_col=5)
    assert not analyze_formula("='#Lookup'!$A$1", source_row=4, source_col=2)
    cross_row = analyze_formula("=C5", source_row=4, source_col=2)
    assert any(issue.code == "cross_row_ref" and issue.is_error for issue in cross_row)
    external = analyze_formula("='[Book.xlsx]Sheet'!$A$1", source_row=4, source_col=2)
    assert any(issue.code == "external_workbook_ref" and issue.is_error for issue in external)
    array = analyze_formula(
        '=TEXTJOIN(",",TRUE,C4:D4)',
        source_row=4,
        source_col=5,
        is_array=True,
    )
    assert any(issue.code == "array_formula" for issue in array)


def test_diff_detects_cell_row_and_column_crud(tmp_path):
    old_path = _write_table(
        tmp_path / "old.xlsx",
        COLUMNS,
        [
            {"Row": "A", "Name": "Alpha", "Power": 1},
            {"Row": "B", "Name": "Beta", "Power": 2},
        ],
    )
    new_path = _write_table(
        tmp_path / "new.xlsx",
        [
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("强度", "Strength", "int32"),
            ("备注", "Note", "FString"),
        ],
        [
            {"Row": "A", "Name": "Alpha2", "Strength": 1, "Note": "changed"},
            {"Row": "C", "Name": "Gamma", "Strength": 3, "Note": "new"},
        ],
    )

    result = diff_files(old_path, new_path)
    sheet_diff = result.sheets[0]
    assert sheet_diff.status == "modified"
    assert any(change.kind == "modified" for change in sheet_diff.column_changes)
    assert any(change.kind == "added" for change in sheet_diff.column_changes)
    assert {change.kind for change in sheet_diff.row_changes} == {"modified", "added", "deleted"}
    text = format_text_diff(result)
    assert "@@ columns @@" in text
    assert "Alpha2" in text
    assert "row C" in text

    output = tmp_path / "changes.xlsx"
    assert Path(write_diff_excel(result, output)).exists()
    wb = openpyxl.load_workbook(output)
    try:
        assert {"Summary", "Changes", "Issues"} <= set(wb.sheetnames)
        assert wb["Changes"].max_row > 1
    finally:
        wb.close()


def test_merge_combines_non_conflicting_cell_row_and_column_changes(tmp_path):
    base = _write_table(
        tmp_path / "base.xlsx",
        COLUMNS,
        [
            {"Row": "A", "Name": "Alpha", "Power": 1},
            {"Row": "B", "Name": "Beta", "Power": 2},
        ],
        blank_middle=True,
    )
    ours = _write_table(
        tmp_path / "ours.xlsx",
        [*COLUMNS, ("辅助", "Helper", "FString")],
        [
            {"Row": "A", "Name": "Alpha", "Power": 10, "Helper": "ours"},
            {"Row": "B", "Name": "Beta", "Power": 2, "Helper": ""},
            {"Row": "C", "Name": "Gamma", "Power": 3, "Helper": "new"},
        ],
    )
    theirs = _write_table(
        tmp_path / "theirs.xlsx",
        COLUMNS,
        [
            {"Row": "Z", "Name": "Before", "Power": 0},
            {"Row": "A", "Name": "Alpha2", "Power": 1},
            {"Row": "B", "Name": "Beta", "Power": 2},
            {"Row": "D", "Name": "Delta", "Power": 4},
        ],
    )

    result = merge_files(base, ours, theirs, strategy="fail")
    assert not result.conflicts
    assert result.ok
    sheet = result.document.sheet("#Data")
    assert [header.en for header in sheet.headers] == ["Row", "Name", "Power", "Helper"]
    by_key = {row.key: row for row in sheet.rows}
    assert set(by_key) == {"A", "B", "C", "Z", "D"}
    assert [row.key for row in sheet.rows] == ["Z", "A", "B", "D", "C"]
    assert by_key["A"].cells[1].value == "Alpha2"
    assert by_key["A"].cells[2].value == 10
    assert by_key["A"].cells[3].value == "ours"

    text = format_merge_text(result)
    assert "conflicts: 0" in text
    output = tmp_path / "merged.xlsx"
    write_merge_excel(result, output)
    headers, rows = _table(output)
    assert headers == ["Row", "Name", "Power", "Helper"]
    assert not any(all(value in (None, "") for value in row) for row in rows)


def test_merge_conflict_can_abort_or_choose_ours(tmp_path):
    base = _write_table(
        tmp_path / "base.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 1}],
    )
    ours = _write_table(
        tmp_path / "ours.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 2}],
    )
    theirs = _write_table(
        tmp_path / "theirs.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 3}],
    )

    failed = merge_files(base, ours, theirs, strategy="fail")
    assert failed.conflicts
    with pytest.raises(MergeConflictError):
        failed.require_resolved()

    chosen = merge_files(base, ours, theirs, strategy="ours")
    assert chosen.conflicts
    assert chosen.document.sheet("#Data").rows[0].cells[2].value == 2
    output = tmp_path / "conflicted.xlsx"
    write_merge_excel(chosen, output, allow_conflicts=True)
    assert Path(output).exists()


def test_delete_vs_modify_conflicts_keep_the_existing_side_for_review(tmp_path):
    base = _write_table(
        tmp_path / "delete_base.xlsx",
        COLUMNS,
        [
            {"Row": "A", "Name": "Alpha", "Power": 1},
            {"Row": "B", "Name": "Beta", "Power": 2},
        ],
    )
    ours = _write_table(
        tmp_path / "delete_ours.xlsx",
        COLUMNS[:2],
        [{"Row": "A", "Name": "Alpha"}],
    )
    theirs = _write_table(
        tmp_path / "delete_theirs.xlsx",
        COLUMNS,
        [
            {"Row": "A", "Name": "Alpha", "Power": 1},
            {"Row": "B", "Name": "Beta changed", "Power": 20},
        ],
    )

    result = merge_files(base, ours, theirs, strategy="fail")
    assert result.conflicts
    assert any(change.kind == "column_delete_conflict" for change in result.conflicts)
    assert any(change.kind == "row_delete_conflict" for change in result.conflicts)
    sheet = result.document.sheet("#Data")
    assert any(header.en == "Power" for header in sheet.headers)
    beta = sheet.get_row("B")
    assert beta.cells[[header.en for header in sheet.headers].index("Name")].value == "Beta changed"


def test_crud_operations_and_formula_rebinding(tmp_path):
    path = tmp_path / "crud.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    for col_index, header in enumerate(
        [
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("A", "A", "FString"),
            ("B", "B", "FString"),
            ("拼接", "Combined", "FString"),
        ],
        start=1,
    ):
        for row_index, value in enumerate(header, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
    ws.cell(row=DATA_START_ROW, column=1, value="A")
    ws.cell(row=DATA_START_ROW, column=2, value="Alpha")
    ws.cell(row=DATA_START_ROW, column=3, value="x")
    ws.cell(row=DATA_START_ROW, column=4, value="y")
    ws.cell(row=DATA_START_ROW, column=5, value="=C4&D4")
    wb.save(path)
    wb.close()

    document = WorkbookDocument.from_xlsx(path)
    apply_operations(
        document,
        [
            {
                "op": "add_column",
                "sheet": "#Data",
                "header": {"cn": "辅助", "en": "Helper", "type": "FString"},
                "position": 2,
                "default": "h",
            },
            {
                "op": "set_cell",
                "sheet": "#Data",
                "row": "A",
                "column": "Name",
                "value": "Alpha2",
            },
            {
                "op": "add_row",
                "sheet": "#Data",
                "row": "B",
                "values": {"Name": "Beta", "A": "p", "B": "q", "Combined": "=D5&E5"},
            },
            {
                "op": "update_column",
                "sheet": "#Data",
                "column": "Helper",
                "en": "Helper2",
            },
        ],
    )
    output = tmp_path / "crud_out.xlsx"
    report = document.validate(clean_first=True)
    document.save(output)
    assert report.ok

    wb = openpyxl.load_workbook(output, data_only=False)
    try:
        ws = wb["#Data"]
        assert [ws.cell(row=2, column=index).value for index in range(1, 7)] == [
            "Row",
            "Name",
            "Helper2",
            "A",
            "B",
            "Combined",
        ]
        assert ws["F4"].value == "=D4&E4"
        assert ws["F5"].value == "=D5&E5"
    finally:
        wb.close()


def test_loading_and_saving_remove_only_fully_blank_rows_and_columns(tmp_path):
    path = tmp_path / "blank.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Data"
    for col_index, values in enumerate(
        [
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("", "", ""),
            ("数值", "Power", "int32"),
        ],
        start=1,
    ):
        for row_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)
    ws.cell(row=4, column=1, value="A")
    ws.cell(row=4, column=2, value="Alpha")
    ws.cell(row=4, column=4, value=1)
    ws.cell(row=5, column=2, value="must be removed with the blank row")
    # Row 6 is intentionally blank.
    ws.cell(row=7, column=1, value="B")
    ws.cell(row=7, column=2, value="Beta")
    ws.cell(row=7, column=4, value=2)
    wb.save(path)
    wb.close()

    document = WorkbookDocument.from_xlsx(path)
    sheet = document.sheet("#Data")
    assert len(sheet.headers) == 3
    assert [row.key for row in sheet.rows] == ["A", "__row_5", "B"]
    output = tmp_path / "blank_out.xlsx"
    document.save(output)
    headers, rows = _table(output)
    assert headers == ["Row", "Name", "Power"]
    assert [row[0] for row in rows] == ["A", "__row_5", "B"]


def test_cli_returns_machine_readable_diff_and_validation_status(tmp_path, capsys):
    old_path = _write_table(
        tmp_path / "cli_old.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 1}],
    )
    new_path = _write_table(
        tmp_path / "cli_new.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 2}],
    )

    assert cli_main(["diff", str(old_path), str(new_path), "--format", "json"]) == 0
    output = capsys.readouterr().out
    payload = json.loads(output)
    assert payload["has_changes"] is True
    assert payload["summary"]["cells_modified"] == 1

    duplicate = _write_table(
        tmp_path / "cli_duplicate.xlsx",
        COLUMNS,
        [
            {"Row": "A", "Name": "Alpha", "Power": 1},
            {"Row": "A", "Name": "Again", "Power": 2},
        ],
    )
    assert cli_main(["validate", str(duplicate), "--json"]) == 1
    output = capsys.readouterr().out
    report = json.loads(output)
    assert report["ok"] is False
    assert any(issue["code"] == "duplicate_row_key" for issue in report["issues"])


def test_cli_edit_applies_json_crud(tmp_path):
    source = _write_table(
        tmp_path / "cli_edit.xlsx",
        COLUMNS,
        [{"Row": "A", "Name": "Alpha", "Power": 1}],
    )
    ops = tmp_path / "ops.json"
    ops.write_text(
        json.dumps(
            [
                {
                    "op": "set_cell",
                    "sheet": "#Data",
                    "row": "A",
                    "column": "Power",
                    "value": 9,
                },
                {
                    "op": "add_row",
                    "sheet": "#Data",
                    "row": "B",
                    "values": {"Name": "Beta", "Power": 2},
                },
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    output = tmp_path / "cli_edit_out.xlsx"

    assert cli_main(
        ["edit", str(source), "--ops", str(ops), "--output", str(output), "--json"]
    ) == 0
    document = WorkbookDocument.from_xlsx(output)
    sheet = document.sheet("#Data")
    assert [row.key for row in sheet.rows] == ["A", "B"]
    assert sheet.get_cell("A", "Power").value == 9
    assert sheet.get_cell("B", "Name").value == "Beta"
