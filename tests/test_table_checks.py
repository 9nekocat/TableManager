# -*- coding: utf-8 -*-
"""Tests for table_checks (extended export validation rules)."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest

import manager_core
import table_checks


def _make_table(path, h1, h2, h3, rows, sheet="#Main"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(h1)
    ws.append(h2)
    ws.append(h3)
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


def _run(tmp_path, monkeypatch, name="T.xlsx"):
    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)
    return manager_core.validate_table(name)


def _by_error(errors, name):
    return [e for e in errors if e["error"] == name]


def test_header_duplicate_english_name(tmp_path, monkeypatch):
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "甲", "乙"],
        ["RowName", "TypeID", "TypeID"],
        ["int", "int", "int"],
        [[1, 10, 20]],
    )
    errors = _run(tmp_path, monkeypatch)
    dup = _by_error(errors, "英文字段名重复")
    assert len(dup) == 1
    assert dup[0]["level"] == "error"
    assert "TypeID" in dup[0]["message"]


def test_header_duplicate_chinese_name_is_warning(tmp_path, monkeypatch):
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "标签", "标签"],
        ["RowName", "TagA", "TagB"],
        ["int", "int", "int"],
        [[1, 10, 20]],
    )
    errors = _run(tmp_path, monkeypatch)
    dup = _by_error(errors, "中文字段名重复")
    assert len(dup) == 1
    assert dup[0]["level"] == "warning"


def test_header_missing_with_data(tmp_path, monkeypatch):
    # 第二列有数据但 H2 为空（None 字面量也视为空）
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "数值"],
        ["RowName", "None"],
        ["int", "int"],
        [[1, 10]],
    )
    errors = _run(tmp_path, monkeypatch)
    miss = _by_error(errors, "表头缺失")
    assert len(miss) == 1
    assert miss[0]["level"] == "warning"
    assert miss[0]["col"] == 2


def test_rowname_empty_and_blank_row(tmp_path, monkeypatch):
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "名称"],
        ["RowName", "Name"],
        ["int", "string"],
        [
            [1, "甲"],
            [None, "有数据但没RowName"],
            [None, None],
            [2, "乙"],
        ],
    )
    errors = _run(tmp_path, monkeypatch)
    no_id = _by_error(errors, "RowName为空")
    assert len(no_id) == 1
    assert no_id[0]["row"] == 5
    assert no_id[0]["level"] == "error"
    blank = _by_error(errors, "空数据行")
    assert len(blank) == 1
    assert blank[0]["row"] == 6
    assert blank[0]["level"] == "warning"


def test_type_mismatch_rules(tmp_path, monkeypatch):
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "整数", "浮点", "布尔", "整数组"],
        ["RowName", "I", "F", "B", "Arr"],
        ["int", "int32", "float", "bool", "array_int"],
        [
            [1, 10, 1.5, True, "1,2,3"],
            [2, "abc", "x", "maybe", "1,x"],
            [3, 1.0, 2, "False", ""],   # 1.0 容忍为 int；空数组跳过
        ],
    )
    errors = _run(tmp_path, monkeypatch)
    bad = _by_error(errors, "类型不匹配")
    # 第 5 行的 4 个错：abc(int)、x(float)、maybe(bool)、1,x(array_int)
    assert len(bad) == 4
    assert {e["row"] for e in bad} == {5}
    assert all(e["level"] == "error" for e in bad)


def test_untyped_and_empty_cells_skipped(tmp_path, monkeypatch):
    _make_table(
        tmp_path / "T.xlsx",
        ["行名", "文本", "未声明"],
        ["RowName", "S", "U"],
        ["int", "string", "None"],
        [[1, "任意文本", "whatever"]],
    )
    errors = _run(tmp_path, monkeypatch)
    assert _by_error(errors, "类型不匹配") == []


def test_missing_files_check(tmp_path, monkeypatch):
    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)
    _make_table(
        tmp_path / "Exists.xlsx",
        ["行名"], ["RowName"], ["int"], [[1]],
    )
    categories = {"categories": {"T": {"files": ["Exists.xlsx", "Ghost.xlsx"]}}}
    errors = manager_core.validate_all(categories)
    miss = _by_error(errors, "配置文件缺失")
    assert len(miss) == 1
    assert miss[0]["file"] == "Ghost.xlsx"
    assert miss[0]["level"] == "error"
