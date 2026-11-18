# -*- coding: utf-8 -*-
"""Tests for ref_index (cross-table reference index)."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest

import manager_core
import ref_index


def _make_table(path, sheets):
    """sheets: {sheet_name: (h1, h2, h3, data_rows)}"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, (h1, h2, h3, rows) in sheets.items():
        ws = wb.create_sheet(name)
        ws.append(h1)
        ws.append(h2)
        ws.append(h3)
        for r in rows:
            ws.append(r)
    wb.save(path)
    wb.close()


@pytest.fixture
def ref_env(tmp_path, monkeypatch):
    """隔离的 RAW_DIR + references.json + categories，清空模块缓存。"""
    _make_table(tmp_path / "Target.xlsx", {
        "#Main": (
            ["行名", "名称"],
            ["RowName", "Name"],
            ["int", "string"],
            [[100, "甲"], [200, "乙"], [300, "丙"]],
        ),
    })
    _make_table(tmp_path / "Source.xlsx", {
        "#Main": (
            ["行名", "目标ID", "描述"],
            ["RowName", "TargetId", "Desc"],
            ["int", "int", "string"],
            [
                [1, 100, "引用存在的 {Buff,200}"],
                [2, 999, "引用不存在的 {Buff,998}"],
                [3, None, "空跳过"],
            ],
        ),
    })

    categories = {"categories": {"Test": {"files": ["Target.xlsx", "Source.xlsx"]}}}
    monkeypatch.setattr(manager_core, "load_categories", lambda: categories)

    ref_cfg = {
        "column_rules": [
            {
                "name": "源-目标",
                "file": "Source.xlsx",
                "column_pattern": "^TargetId$",
                "targets": [{"file": "Target.xlsx"}],
            }
        ],
        "text_patterns": [
            {
                "name": "Buff内嵌引用",
                "regex": "\\{Buff[,，]\\s*(\\d+)\\}",
                "targets": [{"file": "Target.xlsx"}],
            }
        ],
    }
    cfg_path = tmp_path / "references.json"
    cfg_path.write_text(json.dumps(ref_cfg), encoding="utf-8")
    monkeypatch.setattr(ref_index, "CONFIG_PATH", cfg_path)

    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)
    ref_index.load_ref_config(force=True)
    ref_index._space_cache.clear()
    ref_index._index_cache.clear()
    yield tmp_path
    ref_index.load_ref_config(force=True)
    ref_index._space_cache.clear()
    ref_index._index_cache.clear()


def test_column_rule_hit_and_dangling(ref_env):
    index = ref_index.build_reference_index()
    col_records = [r for r in index.records if r["kind"] == "column"]
    assert len(col_records) == 2

    hit = next(r for r in col_records if r["value"] == "100")
    assert not hit["dangling"]
    assert hit["targets_hit"] == [("Target.xlsx", "#Main", 4)]
    assert hit["src_sheet"] == "#Main" and hit["src_row"] == 4

    dangling = next(r for r in col_records if r["value"] == "999")
    assert dangling["dangling"]
    assert dangling in index.dangling


def test_text_pattern_refs(ref_env):
    index = ref_index.build_reference_index()
    text_records = [r for r in index.records if r["kind"] == "text"]
    assert len(text_records) == 2
    by_value = {r["value"]: r for r in text_records}
    assert by_value["200"]["targets_hit"] == [("Target.xlsx", "#Main", 5)]
    assert by_value["998"]["dangling"]


def test_find_references_reverse_lookup(ref_env):
    index = ref_index.build_reference_index()
    # Target.xlsx 第 4 行（ID=100）被 Source.xlsx 的列引用命中
    refs = index.find_references("Target.xlsx", sheet="#Main", row=4)
    assert len(refs) == 1
    assert refs[0]["src_file"] == "Source.xlsx"
    # 不限定行：ID=100（列） + ID=200（文本） 共 2 条
    refs_all = index.find_references("Target.xlsx")
    assert len(refs_all) == 2


def test_effect_type_rules(tmp_path, monkeypatch):
    """Effect 位置参数：Type9 p1->Event, p2->Target, p3+->Effect 自身。"""
    _make_table(tmp_path / "Effect.xlsx", {
        "#Main": (
            ["行名", "效果ID", "效果类型", "参数", "条件ID"],
            ["RowName", "ID", "Type", "Params", "ConditionIDs"],
            ["int", "int", "int", "array_int", "array_int"],
            [
                [500, 500, 9, "10,20,500,999", "30"],
                [501, 501, 3, "1,200,1", ""],
            ],
        ),
    })
    _make_table(tmp_path / "Event.xlsx", {
        "#Main": (["行名"], ["RowName"], ["int"], [[10], [11]]),
    })
    _make_table(tmp_path / "Target.xlsx", {
        "#Main": (["行名"], ["RowName"], ["int"], [[20], [200]]),
    })
    _make_table(tmp_path / "Cond.xlsx", {
        "#Main": (["行名"], ["RowName"], ["int"], [[30]]),
    })

    categories = {"categories": {"T": {"files": ["Effect.xlsx", "Event.xlsx", "Target.xlsx", "Cond.xlsx"]}}}
    monkeypatch.setattr(manager_core, "load_categories", lambda: categories)
    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)

    effect_cfg = {
        "file": "Effect.xlsx",
        "condition_targets": [{"file": "Cond.xlsx"}],
        "types": {
            "9": {
                "name": "事件监听",
                "params": [
                    {"pos": 1, "kind": "ref", "targets": [{"file": "Event.xlsx"}]},
                    {"pos": 2, "kind": "ref", "targets": [{"file": "Target.xlsx"}]},
                    {"pos": "3+", "kind": "ref", "targets": [{"file": "Effect.xlsx"}]},
                ],
            },
            "3": {
                "name": "添加Buff",
                "params": [
                    {"pos": 1, "kind": "literal"},
                    {"pos": 2, "kind": "ref", "targets": [{"file": "Target.xlsx"}]},
                ],
            },
        },
    }
    cfg_path = tmp_path / "effect_checks.json"
    cfg_path.write_text(json.dumps(effect_cfg), encoding="utf-8")
    monkeypatch.setattr(ref_index, "EFFECT_CONFIG_PATH", cfg_path)
    # references.json 指向不存在文件，确保规则只来自 effect_checks.json
    monkeypatch.setattr(ref_index, "CONFIG_PATH", tmp_path / "nonexistent.json")
    ref_index.load_ref_config(force=True)
    ref_index.load_effect_config(force=True)
    ref_index._space_cache.clear()
    ref_index._index_cache.clear()

    try:
        index = ref_index.build_reference_index()
        by_key = {(r["rule"], r["value"]): r for r in index.records}
        # Type9: p1=10->Event 命中；p2=20->Target 命中；p3=500->Effect 自身命中；p4=999 悬空
        assert not by_key[("Effect-Type9-p1", "10")]["dangling"]
        assert not by_key[("Effect-Type9-p2", "20")]["dangling"]
        assert by_key[("Effect-Type9-p3", "500")]["targets_hit"] == [("Effect.xlsx", "#Main", 4)]
        assert by_key[("Effect-Type9-p4", "999")]["dangling"]
        # 条件列（全局 condition_targets）
        assert by_key[("Effect-Type9-cond", "30")]["targets_hit"] == [("Cond.xlsx", "#Main", 4)]
        # Type3: literal 的 p1 不产生记录，仅 p2 检查
        assert ("Effect-Type3-p1", "1") not in by_key
        assert not by_key[("Effect-Type3-p2", "200")]["dangling"]
    finally:
        ref_index.load_ref_config(force=True)
        ref_index.load_effect_config(force=True)
        ref_index._space_cache.clear()
        ref_index._index_cache.clear()


def test_target_column_match(tmp_path, monkeypatch):
    """target.column：按目标表的非 RowName 列（如 Tag）匹配。"""
    _make_table(tmp_path / "Tags.xlsx", {
        "#Stage": (
            ["行名", "标签"],
            ["RowName", "Tag"],
            ["int", "string"],
            [[1, "Stage.Battle.Default"], [2, "Stage.Event.Shop"]],
        ),
    })
    _make_table(tmp_path / "Rooms.xlsx", {
        "#Room": (
            ["行名", "类型Tag"],
            ["RowName", "RawRoomType"],
            ["int", "string"],
            [[1, "Stage.Battle.Default"], [2, "Stage.Unknown"]],
        ),
    })

    categories = {"categories": {"T": {"files": ["Tags.xlsx", "Rooms.xlsx"]}}}
    monkeypatch.setattr(manager_core, "load_categories", lambda: categories)
    monkeypatch.setattr(manager_core, "RAW_DIR", tmp_path)

    ref_cfg = {
        "column_rules": [
            {
                "name": "房间-类型Tag",
                "file": "Rooms.xlsx",
                "column_pattern": "^RawRoomType$",
                "targets": [{"file": "Tags.xlsx", "sheets": ["#Stage"], "column": "Tag"}],
            }
        ]
    }
    cfg_path = tmp_path / "references.json"
    cfg_path.write_text(json.dumps(ref_cfg), encoding="utf-8")
    monkeypatch.setattr(ref_index, "CONFIG_PATH", cfg_path)
    ref_index.load_ref_config(force=True)
    ref_index._space_cache.clear()
    ref_index._index_cache.clear()

    try:
        index = ref_index.build_reference_index()
        records = [r for r in index.records if r["kind"] == "column"]
        assert len(records) == 2
        by_value = {r["value"]: r for r in records}
        assert by_value["Stage.Battle.Default"]["targets_hit"] == [("Tags.xlsx", "#Stage", 4)]
        assert by_value["Stage.Unknown"]["dangling"]
    finally:
        ref_index.load_ref_config(force=True)
        ref_index._space_cache.clear()
        ref_index._index_cache.clear()
