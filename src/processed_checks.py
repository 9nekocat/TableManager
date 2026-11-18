"""
Processed 导出数据（DataTable/Processed/DT_*.json）引用检查。

与 ref_index（Raw xlsx 模式）并存的第二种分析模式：
xlsx 模式用于导表前检查，本模块用于导出后精确分析——导出数据无公式、
无辅助 sheet、无垃圾行，Type 列为自描述枚举名，悬空引用基本零误报。

Processed/*.csv 是陈旧测试数据，本模块只读 DT_*.json。
"""

import ast
import json
from pathlib import Path

import manager_core as mc

CONFIG_PATH = mc.DATA_DIR / "processed_checks.json"
PROCESSED_DIR = mc.PROJECT_ROOT / "DataTable" / "Processed"

# ---------------------------------------------------------------- 配置与加载

_config_cache = None


def load_config(force=False):
    global _config_cache
    if _config_cache is not None and not force:
        return _config_cache
    if CONFIG_PATH.exists():
        try:
            _config_cache = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            _config_cache = {}
    else:
        _config_cache = {}
    return _config_cache


def is_available():
    cfg = load_config()
    return bool(cfg.get("tables")) and PROCESSED_DIR.exists()


def _load_json(path):
    """编码自动探测：utf-8-sig → utf-16 → gbk。失败返回 None。"""
    for enc in ("utf-8-sig", "utf-16", "gbk"):
        try:
            with open(path, encoding=enc) as f:
                return json.load(f)
        except (UnicodeDecodeError, ValueError):
            continue
    return None


def _norm_id(value):
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "none":
        return ""
    try:
        return str(int(float(text)))
    except (ValueError, OverflowError):
        return text


def _to_tokens(value):
    """
    导出值 → token 列表（保持位置，供位置参数按 pos 取）。
    兼容原生 list、字符串 repr（"[1, 2]"）、逗号串、标量、None。
    注意：'0' 普遍作为"无"占位符（无事件/无目标/是否标志位），没有表以 0
    为合法 ID——但不能在这里删除（会破坏位置），由检查方跳过。
    """
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [t for t in (_norm_id(v) for v in value) if t]
    text = str(value).strip()
    if not text or text.lower() == "none":
        return []
    if text.startswith("[") and text.endswith("]"):
        try:
            parsed = ast.literal_eval(text)
            if isinstance(parsed, (list, tuple)):
                return [t for t in (_norm_id(v) for v in parsed) if t]
        except (ValueError, SyntaxError):
            pass
    return [t for t in (_norm_id(p) for p in text.split(",")) if t]


# ---------------------------------------------------------------- 表与 ID 空间

_table_cache = {}


def _load_table(logical):
    """加载逻辑表 → list[dict]（按导出文件 mtime 缓存）。缺失/空表返回 []。"""
    cfg = load_config().get("tables", {}).get(logical)
    if not cfg:
        return []
    path = PROCESSED_DIR / cfg["file"]
    try:
        mtime = path.stat().st_mtime
    except OSError:
        return []
    key = (logical, mtime)
    if key not in _table_cache:
        data = _load_json(path)
        _table_cache.clear()
        _table_cache[key] = data if isinstance(data, list) else []
    return _table_cache[key]


_space_cache = {}


def _target_space(logical, column=None):
    """目标表 ID 空间：{norm_id: row_name}。column 可指定按其他字段匹配（如 AttributeName）。"""
    key = (logical, column)
    if key in _space_cache:
        return _space_cache[key]
    cfg = load_config().get("tables", {}).get(logical, {})
    id_field = column or cfg.get("id_field", "ID")
    space = {}
    for row in _load_table(logical):
        nid = _norm_id(row.get(id_field))
        if nid and nid not in space:
            space[nid] = _norm_id(row.get("Name")) or nid
    _space_cache[key] = space
    return space


def _match_targets(targets, nid):
    hits = []
    for t in targets:
        space = _target_space(t["table"], t.get("column"))
        if nid in space:
            hits.append((t["table"], space[nid]))
    return hits


# ---------------------------------------------------------------- 检查

def _make_record(kind, rule_name, table, row_name, field, value, hits, level):
    return {
        "kind": kind,
        "rule": rule_name,
        "src_file": table,
        "src_sheet": "",
        "src_row": row_name,
        "src_col": 0,
        "src_col_ref": field,
        "src_header": field,
        "value": value,
        "targets_hit": hits,
        "dangling": not hits,
        "level": level,
    }


def _row_name(table, row):
    cfg = load_config().get("tables", {}).get(table, {})
    return _norm_id(row.get(cfg.get("id_field", "ID"))) or _norm_id(row.get("Name"))


def _scan_field_rule(rule, records, skipped):
    table = rule["table"]
    rows = _load_table(table)
    if not rows:
        skipped.append({"rule": rule["name"], "reason": f"表 {table} 无导出数据"})
        return
    present = [f for f in rule["fields"] if any(f in r for r in rows[:50])]
    missing = [f for f in rule["fields"] if f not in present]
    for f in missing:
        skipped.append({"rule": rule["name"], "reason": f"{table} 导出缺少字段 {f}"})
    if not present:
        return
    for row in rows:
        for field in present:
            for nid in _to_tokens(row.get(field)):
                if nid == "0":  # "无"占位符
                    continue
                hits = _match_targets(rule["targets"], nid)
                records.append(_make_record(
                    "processed_field", rule["name"], table, _row_name(table, row),
                    field, nid, hits, rule.get("level", "error"),
                ))


def _scan_positional_rule(rule, records, skipped):
    table = rule["table"]
    rows = _load_table(table)
    if not rows:
        skipped.append({"rule": f"{table} 位置参数", "reason": f"表 {table} 无导出数据"})
        return
    type_field, params_field = rule["type_field"], rule["params_field"]
    if not any(params_field in r for r in rows[:50]):
        skipped.append({"rule": f"{table} 位置参数", "reason": f"{table} 导出缺少字段 {params_field}"})
        return
    types = rule.get("types", {})
    for row in rows:
        type_val = str(row.get(type_field, "")).strip()
        spec = types.get(type_val)
        if not spec:
            continue
        tokens = _to_tokens(row.get(params_field))
        level = spec.get("level", "error")
        enum_short = type_val.split("::")[-1]
        param_offset = 1 if table == "Effect" and enum_short == "Chance" and tokens and tokens[0] == "PositiveEventValue" else 0
        for p in spec.get("params", []):
            pos = p["pos"]
            if isinstance(pos, str) and pos.endswith("+"):
                indices = range(int(pos[:-1]) - 1 + param_offset, len(tokens))
            else:
                indices = [int(pos) - 1 + param_offset] if int(pos) - 1 + param_offset < len(tokens) else []
            for i in indices:
                nid = tokens[i] if i < len(tokens) else ""
                if not nid or nid == "0":  # 空 / "无"占位符
                    continue
                hits = _match_targets(p["targets"], nid)
                records.append(_make_record(
                    "processed_positional", f"{table}-{enum_short}-p{i + 1}", table,
                    _row_name(table, row), params_field, nid, hits, level,
                ))


_result_cache = {}


def run_checks(table_filter=None, use_cache=True):
    """
    运行全部规则，返回 {records, dangling, skipped, stats}。
    table_filter: 逻辑表名，只检查该表为来源的记录。
    """
    cfg = load_config()
    files_sig = []
    for logical, t in cfg.get("tables", {}).items():
        path = PROCESSED_DIR / t["file"]
        try:
            files_sig.append((logical, path.stat().st_mtime))
        except OSError:
            files_sig.append((logical, None))
    sig = (tuple(sorted(files_sig)), table_filter or "")
    if use_cache and _result_cache.get("sig") == sig:
        return _result_cache["result"]

    records, skipped = [], []
    for rule in cfg.get("field_rules", []):
        if table_filter and rule["table"] != table_filter:
            continue
        try:
            _scan_field_rule(rule, records, skipped)
        except Exception as e:
            skipped.append({"rule": rule.get("name", "?"), "reason": f"规则执行失败: {e}"})
    for rule in cfg.get("positional_rules", []):
        if table_filter and rule["table"] != table_filter:
            continue
        try:
            _scan_positional_rule(rule, records, skipped)
        except Exception as e:
            skipped.append({"rule": f"{rule.get('table', '?')} 位置参数", "reason": f"规则执行失败: {e}"})

    dangling = [r for r in records if r["dangling"]]
    by_rule = {}
    for r in dangling:
        by_rule[r["rule"]] = by_rule.get(r["rule"], 0) + 1
    result = {
        "records": records,
        "dangling": dangling,
        "skipped": skipped,
        "stats": {
            "total_records": len(records),
            "total_dangling": len(dangling),
            "dangling_by_rule": dict(sorted(by_rule.items(), key=lambda kv: -kv[1])),
        },
    }
    _result_cache.clear()
    _result_cache.update({"sig": sig, "result": result})
    return result
