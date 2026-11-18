# -*- coding: utf-8 -*-
"""Tests for processed_checks (Processed/DT_*.json export analysis)."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pytest

import processed_checks as pc


def _write_json(path, rows, encoding="utf-8-sig"):
    path.write_text(json.dumps(rows, ensure_ascii=False), encoding=encoding)


@pytest.fixture
def p_env(tmp_path, monkeypatch):
    """隔离的 Processed 目录 + 规则配置，清空模块缓存。"""
    pdir = tmp_path / "Processed"
    pdir.mkdir()

    _write_json(pdir / "DT_Target.json", [
        {"TargetId": "100", "Name": "100"},
        {"TargetId": "200", "Name": "200"},
    ])
    _write_json(pdir / "DT_Buff.json", [
        {"BuffID": "1001", "Name": "1001"},
    ], encoding="utf-16")  # 验证 utf-16 探测
    _write_json(pdir / "DT_Attr.json", [
        {"AttributeId": "10100010", "AttributeName": "MaxHealth", "Name": "10100010"},
    ])
    _write_json(pdir / "DT_Card.json", [
        # Tags 原生数组；EffectIds 字符串 repr；含 0 占位符与悬空值
        {"CardId": "100010", "Tags": [1, 2], "EffectIds": "['9000001', '0']", "Name": "100010"},
        {"CardId": "100020", "Tags": [999], "EffectIds": None, "Name": "100020"},
    ])
    _write_json(pdir / "DT_Effect.json", [
        {"ID": "9000001", "Type": "ECardEffectType::Buff", "Params": ['1', '1001', '1'],
         "ConditionIDs": None, "Name": "9000001"},
        # p1=0 占位：跳过且不得影响后续位置
        {"ID": "9000002", "Type": "ECardEffectType::Event", "Params": ['0', '100', '9000001'],
         "ConditionIDs": None, "Name": "9000002"},
        {"ID": "9000003", "Type": "ECardEffectType::Event", "Params": ['5', '999', '123456'],
         "ConditionIDs": None, "Name": "9000003"},
        {"ID": "9000004", "Type": "ECardEffectType::Damage", "Params": ['1', '999'],
         "ConditionIDs": None, "Name": "9000004"},
        {"ID": "9000005", "Type": "ECardEffectType::ModifyAttribute", "Params": ['MaxHealth', '1'],
         "ConditionIDs": None, "Name": "9000005"},
        {"ID": "9000006", "Type": "ECardEffectType::ModifyAttribute", "Params": ['NoSuchAttr', '1'],
         "ConditionIDs": None, "Name": "9000006"},
    ])
    _write_json(pdir / "DT_Event.json", [
        {"ID": "5", "SourceID": "100", "Type": "ERavenEventType::OnCoolDown",
         "Params": None, "Name": "5"},
    ])
    _write_json(pdir / "DT_Tag.json", [
        {"TagId": 1, "Name": "1"}, {"TagId": 2, "Name": "2"},
    ])

    cfg = {
        "tables": {
            "Card": {"file": "DT_Card.json", "id_field": "CardId"},
            "Effect": {"file": "DT_Effect.json", "id_field": "ID"},
            "Event": {"file": "DT_Event.json", "id_field": "ID"},
            "Target": {"file": "DT_Target.json", "id_field": "TargetId"},
            "Buff": {"file": "DT_Buff.json", "id_field": "BuffID"},
            "Tag": {"file": "DT_Tag.json", "id_field": "TagId"},
            "Attribute": {"file": "DT_Attr.json", "id_field": "AttributeId"},
            "Ghost": {"file": "DT_Ghost.json", "id_field": "ID"},
        },
        "field_rules": [
            {"name": "卡牌-标签", "table": "Card", "fields": ["Tags"],
             "targets": [{"table": "Tag"}]},
            {"name": "卡牌-效果", "table": "Card", "fields": ["EffectIds"],
             "targets": [{"table": "Effect"}]},
            {"name": "事件-目标", "table": "Event", "fields": ["SourceID"],
             "targets": [{"table": "Target"}]},
            {"name": "幽灵规则", "table": "Ghost", "fields": ["X"],
             "targets": [{"table": "Tag"}]},
            {"name": "缺字段规则", "table": "Card", "fields": ["NoSuchField"],
             "targets": [{"table": "Tag"}]},
        ],
        "positional_rules": [
            {
                "table": "Effect",
                "type_field": "Type",
                "params_field": "Params",
                "types": {
                    "ECardEffectType::Buff": {
                        "params": [{"pos": 2, "targets": [{"table": "Buff"}]}],
                    },
                    "ECardEffectType::Event": {
                        "params": [
                            {"pos": 1, "targets": [{"table": "Event"}]},
                            {"pos": 2, "targets": [{"table": "Target"}]},
                            {"pos": "3+", "targets": [{"table": "Effect"}]},
                        ],
                    },
                    "ECardEffectType::ModifyAttribute": {
                        "level": "warning",
                        "params": [{"pos": 1, "targets": [{"table": "Attribute", "column": "AttributeName"}]}],
                    },
                },
            }
        ],
    }
    cfg_path = tmp_path / "processed_checks.json"
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")

    monkeypatch.setattr(pc, "CONFIG_PATH", cfg_path)
    monkeypatch.setattr(pc, "PROCESSED_DIR", pdir)
    pc.load_config(force=True)
    pc._table_cache.clear()
    pc._space_cache.clear()
    pc._result_cache.clear()
    yield pdir
    pc.load_config(force=True)
    pc._table_cache.clear()
    pc._space_cache.clear()
    pc._result_cache.clear()


def test_field_rules_hit_dangling_and_zero(p_env):
    res = pc.run_checks()
    recs = {(r["rule"], r["src_row"], r["value"]): r for r in res["records"]}
    # 原生数组命中
    assert not recs[("卡牌-标签", "100010", "1")]["dangling"]
    # 悬空
    assert recs[("卡牌-标签", "100020", "999")]["dangling"]
    # 字符串 repr 数组解析 + 命中
    assert not recs[("卡牌-效果", "100010", "9000001")]["dangling"]
    # '0' 占位符不产生记录
    assert ("卡牌-效果", "100010", "0") not in recs
    # 目标命中行名
    assert recs[("事件-目标", "5", "100")]["targets_hit"] == [("Target", "100")]


def test_positional_rules(p_env):
    res = pc.run_checks()
    recs = {(r["rule"], r["src_row"], r["value"]): r for r in res["records"]}
    # Buff p2 命中（utf-16 加载的表）
    assert not recs[("Effect-Buff-p2", "9000001", "1001")]["dangling"]
    # Event: p1=0 跳过且位置不偏移——p2 仍取到 100 命中 Target
    assert not recs[("Effect-Event-p2", "9000002", "100")]["dangling"]
    # "3+" 自引用命中
    assert recs[("Effect-Event-p3", "9000002", "9000001")]["targets_hit"] == [("Effect", "9000001")]
    # 悬空：事件 999 / 效果 123456
    assert recs[("Effect-Event-p2", "9000003", "999")]["dangling"]
    assert recs[("Effect-Event-p3", "9000003", "123456")]["dangling"]
    # Damage 无规则，不产生记录
    assert not any(r["src_row"] == "9000004" for r in res["records"] if r["kind"] == "processed_positional")
    # target.column 匹配 + warning 级别
    attr_hit = recs[("Effect-ModifyAttribute-p1", "9000005", "MaxHealth")]
    assert not attr_hit["dangling"] and attr_hit["level"] == "warning"
    assert recs[("Effect-ModifyAttribute-p1", "9000006", "NoSuchAttr")]["dangling"]


def test_missing_table_and_field_skipped(p_env):
    res = pc.run_checks()
    reasons = {s["rule"]: s["reason"] for s in res["skipped"]}
    assert "幽灵规则" in reasons
    assert "缺字段规则" in reasons
    # 跳过的规则不产生记录也不报错
    assert not any(r["rule"] in ("幽灵规则", "缺字段规则") for r in res["records"])


def test_table_filter(p_env):
    res = pc.run_checks(table_filter="Card")
    assert {r["src_file"] for r in res["records"]} == {"Card"}
