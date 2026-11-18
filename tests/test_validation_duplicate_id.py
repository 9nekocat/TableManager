# -*- coding: utf-8 -*-
"""Validation tests for duplicate RowName IDs."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl

import manager_core


def _create_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Main"
    ws.append(["行名", "名称", "数值"])
    ws.append(["RowName", "Name", "Value"])
    ws.append(["FString", "FString", "int32"])
    ws.append(["ID_A", "Alpha", 1])
    ws.append(["ID_B", "Beta", 2])
    ws.append(["ID_A", "Gamma", 3])
    ws.append(["", "Blank row name is ignored", 4])

    other = wb.create_sheet("#Other")
    other.append(["行名", "名称"])
    other.append(["RowName", "Name"])
    other.append(["FString", "FString"])
    other.append(["ID_A", "Allowed in another sheet"])

    wb.save(path)
    wb.close()


def test_validate_table_reports_duplicate_rowname_ids(tmp_path, monkeypatch):
    filepath = tmp_path / "DuplicateIds.xlsx"
    _create_workbook(filepath)
    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)

    errors = manager_core.validate_table("DuplicateIds.xlsx")
    duplicate_errors = [
        e for e in errors
        if e["error"] == manager_core.DUPLICATE_ID_ERROR
    ]

    assert len(duplicate_errors) == 2
    assert {e["sheet"] for e in duplicate_errors} == {"#Main"}
    assert {e["row"] for e in duplicate_errors} == {4, 6}
    assert {e["col_ref"] for e in duplicate_errors} == {"A4", "A6"}
    assert all(e["value"] == "ID_A" for e in duplicate_errors)
    assert all(e["duplicate_rows"] == [4, 6] for e in duplicate_errors)
