"""
跨表引用索引。

按 references.json 的规则扫描 RAW_DIR 下的配置表，产出：
  - 引用记录列表（列级 ID 引用 / 文本内嵌引用 / Effect 位置参数引用 / 跨sheet与外链公式引用）
  - 反向索引：目标 (file, sheet, rowname) -> 引用者列表
  - 悬空引用列表（目标不存在）

供 table_checks（导表检查）、CLI `refs` 子命令和 `/api/refs` 端点共用。

性能注意：openpyxl 的 read_only 模式不支持 ws.cell() 随机访问（每次都会
重新流式解析，大表 O(n²)），因此本模块统一使用 eager 加载。
"""

import json
import re
from pathlib import Path

import openpyxl

import manager_core as mc
from formula_utils import parse_cell_refs
from merge_engine import _external_link_index_map

CONFIG_PATH = mc.DATA_DIR / "references.json"
EFFECT_CONFIG_PATH = mc.DATA_DIR / "effect_checks.json"

# ---------------------------------------------------------------- 配置加载

_config_cache = None
_effect_config_cache = None


def load_ref_config(force=False):
    """加载 references.json；文件不存在时返回空配置（引用检查整体跳过）。"""
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


def load_effect_config(force=False):
    """加载 effect_checks.json（Effect.xlsx 按 Type 的位置参数规则）。"""
    global _effect_config_cache
    if _effect_config_cache is not None and not force:
        return _effect_config_cache
    if EFFECT_CONFIG_PATH.exists():
        try:
            _effect_config_cache = json.loads(EFFECT_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            _effect_config_cache = {}
    else:
        _effect_config_cache = {}
    return _effect_config_cache


def effect_type_rules():
    """
    把 effect_checks.json 转换为内部规则列表：
    [{type, level?, params: [{pos, targets}], condition_ids_targets?}]
    仅 kind == "ref" 且配置了 targets 的参数位置参与检查（literal 仅作文档）。
    """
    cfg = load_effect_config()
    cond_targets = cfg.get("condition_targets")
    rules = []
    for type_str, spec in cfg.get("types", {}).items():
        params = [
            {"pos": p["pos"], "targets": p["targets"]}
            for p in spec.get("params", [])
            if p.get("kind") == "ref" and p.get("targets")
        ]
        rule = {"type": int(type_str), "params": params}
        if spec.get("level"):
            rule["level"] = spec["level"]
        if cond_targets:
            rule["condition_ids_targets"] = cond_targets
        if params or cond_targets:
            rules.append(rule)
    return rules


def is_available():
    """references.json / effect_checks.json 配置了任何规则时，引用功能才生效。"""
    cfg = load_ref_config()
    return bool(
        cfg.get("column_rules") or cfg.get("text_patterns") or effect_type_rules()
    )


# ---------------------------------------------------------------- 工具

def _norm_id(value):
    """归一化 ID：数值统一为整数字符串，其余为 strip 后文本；空返回 ''。"""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        return str(int(float(text)))
    except (ValueError, OverflowError):
        return text


def _is_empty(value):
    return value is None or str(value).strip() == ""


def _iter_config_files(file_filter=None):
    cats = mc.load_categories()
    files = []
    for cat in cats.get("categories", {}).values():
        files.extend(cat.get("files", []))
    if cats.get("uncategorized"):
        files.extend(cats["uncategorized"].get("files", []))
    seen, result = set(), []
    for rel in files:
        if rel in seen:
            continue
        seen.add(rel)
        if file_filter and rel != file_filter:
            continue
        if (mc.RAW_DIR / rel).exists():
            result.append(rel)
    return result


def _load_wb(relpath, data_only=True):
    """eager 加载（不用 read_only：随机单元格访问在 read_only 下极慢）。"""
    try:
        return openpyxl.load_workbook(mc.RAW_DIR / relpath, data_only=data_only)
    except Exception:
        return None


# ---------------------------------------------------------------- 目标 ID 空间

_space_cache = {}


def _target_space(relpath, sheets=None, column=None):
    """
    目标表的 ID 空间：{norm_id: (sheet_name, row_idx)}。
    column 为 None 时取 A 列（RowName），否则按表头名（第1/2行）定位列。
    按 (file, sheets, column, mtime) 缓存。
    """
    full = mc.RAW_DIR / relpath
    try:
        mtime = full.stat().st_mtime
    except OSError:
        return {}
    key = (relpath, tuple(sheets or ()), column, mtime)
    if key in _space_cache:
        return _space_cache[key]

    space = {}
    wb = _load_wb(relpath)
    if wb:
        for sn in wb.sheetnames:
            if not sn.startswith("#"):
                continue
            if sheets and sn not in sheets:
                continue
            ws = wb[sn]
            col_idx = 1
            if column:
                col_idx = None
                for c in range(1, ws.max_column + 1):
                    h1 = ws.cell(row=1, column=c).value
                    h2 = ws.cell(row=2, column=c).value
                    if column in (str(h1).strip() if h1 else "", str(h2).strip() if h2 else ""):
                        col_idx = c
                        break
                if col_idx is None:
                    continue
            for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
                nid = _norm_id(ws.cell(row=row_idx, column=col_idx).value)
                if nid and nid not in space:
                    space[nid] = (sn, row_idx)
        wb.close()
    _space_cache[key] = space
    return space


def _match_targets(targets, nid):
    """在候选目标列表中查找 nid，返回 [(file, sheet, row), ...]（可能多个候选命中）。"""
    hits = []
    for t in targets:
        space = _target_space(t["file"], t.get("sheets"), t.get("column"))
        if nid in space:
            sn, row_idx = space[nid]
            hits.append((t["file"], sn, row_idx))
    return hits


# ---------------------------------------------------------------- 扫描器

def _make_record(kind, rule_name, src_file, src_sheet, row, col, header, value, hits, level="error"):
    return {
        "kind": kind,
        "rule": rule_name,
        "src_file": src_file,
        "src_sheet": src_sheet,
        "src_row": row,
        "src_col": col,
        "src_col_ref": f"{openpyxl.utils.get_column_letter(col)}{row}",
        "src_header": header,
        "value": value,
        "targets_hit": hits,
        "dangling": not hits,
        "level": level,
    }


def _scan_column_rule(rule, records):
    src_file = rule["file"]
    pattern = re.compile(rule["column_pattern"])
    sep = rule.get("array_sep", ",")
    sheets_filter = rule.get("sheets")
    wb = _load_wb(src_file)
    if not wb:
        return
    for sn in wb.sheetnames:
        if not sn.startswith("#"):
            continue
        if sheets_filter and sn not in sheets_filter:
            continue
        ws = wb[sn]
        # 列匹配：同时对照第1行（中文）和第2行（英文）表头
        matched_cols = []
        for c in range(1, ws.max_column + 1):
            h1 = ws.cell(row=1, column=c).value
            h2 = ws.cell(row=2, column=c).value
            header = ""
            for h in (h2, h1):
                text = str(h).strip() if h is not None else ""
                if text and text != "None":
                    header = text
                    break
            if pattern.match(str(h1).strip() if h1 else "") or pattern.match(
                str(h2).strip() if h2 else ""
            ):
                matched_cols.append((c, header))
        if not matched_cols:
            continue
        for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
            for c, header in matched_cols:
                raw = ws.cell(row=row_idx, column=c).value
                if _is_empty(raw):
                    continue
                for token in str(raw).split(sep):
                    nid = _norm_id(token)
                    if not nid:
                        continue
                    hits = _match_targets(rule["targets"], nid)
                    records.append(
                        _make_record("column", rule["name"], src_file, sn, row_idx, c, header, nid, hits)
                    )
    wb.close()


def _scan_text_patterns(rules, files, records):
    """所有文本内嵌模式合并为单遍扫描（避免每个正则重复打开全部文件）。"""
    compiled = [(re.compile(r["regex"]), r) for r in rules]
    for src_file in files:
        wb = _load_wb(src_file)
        if not wb:
            continue
        for sn in wb.sheetnames:
            if not sn.startswith("#"):
                continue
            ws = wb[sn]
            for row in ws.iter_rows(min_row=mc.DATA_START_ROW):
                for cell in row:
                    v = cell.value
                    if not isinstance(v, str) or "{" not in v:
                        continue
                    for regex, rule in compiled:
                        for m in regex.finditer(v):
                            nid = _norm_id(m.group(1))
                            if not nid:
                                continue
                            hits = _match_targets(rule["targets"], nid)
                            records.append(
                                _make_record(
                                    "text", rule["name"], src_file, sn, cell.row, cell.column,
                                    f"{{{m.group(0)}}}", nid, hits,
                                )
                            )
        wb.close()


def _scan_effect_type_rules(rules, records):
    """Effect 表专用：按 Type 列决定 Params 各位置的引用语义（规则见 effect_checks.json）。"""
    src_file = load_effect_config().get("file", "Effect.xlsx")
    wb = _load_wb(src_file)
    if not wb:
        return
    rules_by_type = {int(r["type"]): r for r in rules}
    for sn in wb.sheetnames:
        if not sn.startswith("#"):
            continue
        ws = wb[sn]
        h2 = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=2, column=c).value
            if v is not None:
                h2[str(v).strip()] = c
        type_col, params_col = h2.get("Type"), h2.get("Params")
        if not type_col or not params_col:
            continue
        cond_col = h2.get("ConditionIDs")
        for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
            try:
                row_type = int(float(str(ws.cell(row=row_idx, column=type_col).value)))
            except (TypeError, ValueError):
                continue
            rule = rules_by_type.get(row_type)
            if not rule:
                continue
            level = rule.get("level", "error")
            raw_params = ws.cell(row=row_idx, column=params_col).value
            tokens = [t.strip() for t in str(raw_params).split(",")] if not _is_empty(raw_params) else []
            param_offset = 1 if row_type == 8 and tokens and tokens[0] == "PositiveEventValue" else 0
            for p in rule.get("params", []):
                pos = p["pos"]
                if isinstance(pos, str) and pos.endswith("+"):
                    indices = range(int(pos[:-1]) - 1 + param_offset, len(tokens))
                else:
                    indices = [int(pos) - 1 + param_offset] if int(pos) - 1 + param_offset < len(tokens) else []
                for i in indices:
                    nid = _norm_id(tokens[i]) if i < len(tokens) else ""
                    if not nid:
                        continue
                    hits = _match_targets(p["targets"], nid)
                    records.append(
                        _make_record(
                            "effect_params", f"Effect-Type{row_type}-p{i + 1}",
                            src_file, sn, row_idx, params_col, "Params", nid, hits, level,
                        )
                    )
            if cond_col and rule.get("condition_ids_targets"):
                raw_cond = ws.cell(row=row_idx, column=cond_col).value
                if not _is_empty(raw_cond):
                    for token in str(raw_cond).split(","):
                        nid = _norm_id(token)
                        if not nid:
                            continue
                        hits = _match_targets(rule["condition_ids_targets"], nid)
                        records.append(
                            _make_record(
                                "effect_condition", f"Effect-Type{row_type}-cond",
                                src_file, sn, row_idx, cond_col, "ConditionIDs", nid, hits, level,
                            )
                        )
    wb.close()


def _scan_formulas(files, basename_to_rel, records):
    """跨 sheet / 跨文件公式引用（表内拼接公式属于噪音，跳过）。"""
    for src_file in files:
        wb = _load_wb(src_file, data_only=False)
        if not wb:
            continue
        wb_vals = None  # 惰性打开（仅当确有跨表引用时）
        ext_slots = None  # 惰性解析外链槽位
        for sn in wb.sheetnames:
            if not sn.startswith("#"):
                continue
            ws = wb[sn]
            for row in ws.iter_rows(min_row=mc.DATA_START_ROW):
                for cell in row:
                    v = cell.value
                    # 跨表引用必含 '!'，表内拼接公式直接跳过
                    if not isinstance(v, str) or not v.startswith("=") or "!" not in v:
                        continue
                    for ref in parse_cell_refs(v):
                        target_file, target_sheet = None, None
                        if ref.get("file"):
                            slot_m = re.fullmatch(r"\[?(\d+)\]?", ref["file"])
                            if slot_m:
                                if ext_slots is None:
                                    ext_slots = _external_link_index_map(mc.RAW_DIR / src_file)
                                stem = ext_slots.get(slot_m.group(1))
                                target_file = basename_to_rel.get(stem.lower()) if stem else None
                            else:
                                target_file = basename_to_rel.get(ref["file"].lower())
                            target_sheet = ref.get("sheet")
                        elif ref.get("sheet") and ref["sheet"] != sn:
                            target_file, target_sheet = src_file, ref["sheet"]
                        if not target_file or not target_sheet:
                            continue
                        # 只把指向数据 sheet 数据区的公式视为数据引用；
                        # 指向映射/说明等辅助 sheet 或表头行的是辅助公式，跳过
                        row_idx = ref.get("row_start")
                        if not target_sheet.startswith("#") or not row_idx or row_idx < mc.DATA_START_ROW:
                            continue
                        # 解析目标行：行存在且 A 列有 RowName 即合法
                        hits = []
                        if wb_vals is None:
                            wb_vals = _load_wb(src_file, data_only=True)
                        t_wb, owns = (wb_vals, False) if target_file == src_file else (_load_wb(target_file), True)
                        if t_wb and target_sheet in t_wb.sheetnames:
                            t_ws = t_wb[target_sheet]
                            if row_idx <= t_ws.max_row:
                                nid = _norm_id(t_ws.cell(row=row_idx, column=1).value)
                                if nid:
                                    hits.append((target_file, target_sheet, row_idx))
                        if owns and t_wb:
                            t_wb.close()
                        display = (
                            f"[{ref['file']}]{target_sheet}!{ref['cell_ref']}"
                            if ref.get("file")
                            else f"{target_sheet}!{ref['cell_ref']}"
                        )
                        records.append(
                            _make_record(
                                "formula", "公式引用", src_file, sn, cell.row, cell.column,
                                "", display, hits, "warning",
                            )
                        )
        wb.close()
        if wb_vals:
            wb_vals.close()


# ---------------------------------------------------------------- 索引构建与查询

class RefIndex:
    def __init__(self, records):
        self.records = records
        self.by_target = {}
        self.dangling = []
        for rec in records:
            if rec["dangling"]:
                self.dangling.append(rec)
            for tfile, tsheet, trow in rec["targets_hit"]:
                self.by_target.setdefault((tfile, tsheet, trow), []).append(rec)

    def find_references(self, file, sheet=None, row=None):
        """查找引用指定表（可选限定 sheet/行号）的所有来源记录。"""
        result = []
        for (tfile, tsheet, trow), recs in self.by_target.items():
            if tfile != file:
                continue
            if sheet and tsheet != sheet:
                continue
            if row is not None and trow != row:
                continue
            result.extend(recs)
        return result


_index_cache = {}


def build_reference_index(file_filter=None, use_cache=True):
    """
    构建引用索引。file_filter 限定扫描来源文件（目标空间仍覆盖全部表）。
    按全部涉及文件的 mtime 签名缓存。
    """
    cfg = load_ref_config()
    files = _iter_config_files(file_filter)
    all_files = _iter_config_files()
    sig = tuple(
        sorted((f, (mc.RAW_DIR / f).stat().st_mtime) for f in all_files if (mc.RAW_DIR / f).exists())
    ) + (file_filter or "",)
    if use_cache and _index_cache.get("sig") == sig:
        return _index_cache["index"]

    records = []
    basename_to_rel = {Path(f).stem.lower(): f for f in all_files}

    for rule in cfg.get("column_rules", []):
        if file_filter and rule.get("file") != file_filter:
            continue
        try:
            _scan_column_rule(rule, records)
        except Exception:
            continue

    text_rules = cfg.get("text_patterns", [])
    if text_rules:
        try:
            _scan_text_patterns(text_rules, files, records)
        except Exception:
            pass

    effect_rules = effect_type_rules()
    effect_file = load_effect_config().get("file", "Effect.xlsx")
    if effect_rules and (not file_filter or file_filter == effect_file):
        try:
            _scan_effect_type_rules(effect_rules, records)
        except Exception:
            pass

    try:
        _scan_formulas(files, basename_to_rel, records)
    except Exception:
        pass

    index = RefIndex(records)
    _index_cache.clear()
    _index_cache.update({"sig": sig, "index": index})
    return index
