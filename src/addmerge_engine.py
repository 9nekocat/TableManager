# -*- coding: utf-8 -*-
"""
Add-Merge（增量合并）引擎

将「暂存表」（staging xlsx）合并进 DataTable/Raw 的目标表：
- 只接受新增行与已有行的修改，**绝不删除目标表的任何行**
- 行匹配按 RowName（A 列计算值，无缓存回退 B 列，与 merge_engine 一致）
- 列对齐按第 2 行英文字段名（复用 merge_engine 的语义列对齐）
- staging 为空的单元格不覆盖 target（防止误删数据）
- apply 时走事务：备份 → openpyxl 写回 → 外链/动态数组元数据恢复 →
  Excel COM 重算公式缓存；任一步失败回滚备份
"""

from __future__ import annotations

import os
import sys
import tempfile
import zipfile
from collections import OrderedDict
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from manager_core import DATA_START_ROW
from merge_engine import (
    _read_all_rows_with_formulas,
    _header_signatures,
    _canonical_columns,
    align_datasets_by_headers,
    _read_external_link_cache,
    _inject_external_links,
    _inject_dynamic_array_metadata,
)
from merge_workflow import create_merge_backup, restore_merge_backup
from formula_utils import adjust_formula, normalize_formula_for_compare


class AddMergeError(RuntimeError):
    """增量合并失败（写回前校验失败，或 apply 事务失败已回滚）。"""

    def __init__(self, message, *, backup_path=None, restored=False):
        super().__init__(message)
        self.backup_path = str(backup_path) if backup_path else None
        self.restored = restored


# ---------------------------------------------------------------------------
# 单元格比较工具
# ---------------------------------------------------------------------------

def _is_empty_cell(value, formula):
    """staging 单元格为空：无公式且值为 None/空白字符串。"""
    if formula:
        return False
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _cell_equal(a, b):
    """宽松比较两个计算值：None 与空串等价，数字 1 与 1.0 等价。"""
    if a is None:
        a = ""
    if b is None:
        b = ""
    if isinstance(a, str):
        a = a.strip()
    if isinstance(b, str):
        b = b.strip()
    if a == b:
        return True
    # 类型不同（如 int 1 vs str "1"、float 1.0）时按数值/字符串兜底
    try:
        if float(a) == float(b):
            return True
    except (TypeError, ValueError):
        pass
    return str(a) == str(b)


def _disp(value, formula):
    """报告用单元格显示文本。"""
    if formula:
        if value is not None and str(value).strip() != "":
            return f"{formula} → {value}"
        return str(formula)
    return "" if value is None else str(value)


# ---------------------------------------------------------------------------
# RowName 重复校验（merge_engine 读取时后者覆盖前者，需独立扫描）
# ---------------------------------------------------------------------------

def _find_duplicate_rownames(path, sheets=None):
    """扫描 # 开头 sheet 的 RowName（A 列计算值，无缓存回退 B 列）。

    返回: [{"sheet": sn, "rowname": rn, "count": n}, ...]
    """
    dups = []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for sn in wb.sheetnames:
            if not sn.startswith("#"):
                continue
            if sheets is not None and sn not in sheets:
                continue
            counter = {}
            for row in wb[sn].iter_rows(min_row=DATA_START_ROW, values_only=True):
                if not row:
                    continue
                rn = row[0]
                if rn is None or str(rn).strip() == "":
                    rn = row[1] if len(row) > 1 else None
                if rn is None or str(rn).strip() == "":
                    continue
                key = str(rn).strip()
                counter[key] = counter.get(key, 0) + 1
            for key, cnt in counter.items():
                if cnt > 1:
                    dups.append({"sheet": sn, "rowname": key, "count": cnt})
    finally:
        wb.close()
    return dups


# ---------------------------------------------------------------------------
# 合并计划（dry-run 与 apply 共用）
# ---------------------------------------------------------------------------

def _match_target_sheet(staging_sheet, target_sheetnames):
    """sheet 名大小写不敏感匹配，返回 target 实际 sheet 名或 None。"""
    lowered = staging_sheet.lower()
    for sn in target_sheetnames:
        if sn.lower() == lowered:
            return sn
    return None


def _row_of(sheet_origins, rowname):
    """从 origins 取行号（A 列坐标）。"""
    row_origins = sheet_origins.get(rowname) or {}
    origin = row_origins.get(1)
    return origin[0] if origin else None


def _plan_sheet(stg_sn, tgt_sn, stg_data, stg_headers, stg_origins,
                tgt_data, tgt_headers, tgt_origins):
    """计算单个 sheet 的合并计划。

    返回 (sheet_report, sheet_plan)。sheet_plan 供写回阶段使用。
    """
    stg_hdr = stg_headers.get(stg_sn, {})
    tgt_hdr = tgt_headers.get(tgt_sn, {})
    stg_rows = stg_data.get(stg_sn, OrderedDict())
    tgt_rows = tgt_data.get(tgt_sn, OrderedDict())
    stg_sheet_origins = stg_origins.get(stg_sn, {})
    tgt_sheet_origins = tgt_origins.get(tgt_sn, {})

    # 语义列对齐：canonical 列序以 target 为准，staging 独有列追加在尾部
    columns = _canonical_columns(tgt_hdr, stg_hdr)
    aligned_tgt, aligned_stg, _ = align_datasets_by_headers(
        {tgt_sn: tgt_rows}, {tgt_sn: tgt_hdr},
        {tgt_sn: stg_rows}, {tgt_sn: stg_hdr},
    )
    aligned_tgt_rows = aligned_tgt.get(tgt_sn, OrderedDict())
    aligned_stg_rows = aligned_stg.get(tgt_sn, OrderedDict())

    # canonical 列 → 两侧物理列
    tgt_sig_to_phys = {m["sig"]: phys for phys, m in _header_signatures(tgt_hdr).items()}
    stg_sig_to_phys = {m["sig"]: phys for phys, m in _header_signatures(stg_hdr).items()}
    stg_phys_to_meta = _header_signatures(stg_hdr)

    report = {
        "staging_sheet": stg_sn,
        "target_sheet": tgt_sn,
        "status": "ok",
        "skip_reason": None,
        "added": [],
        "modified": [],
        "unchanged": [],
        "skipped_columns": [],
    }
    plan = {"staging_sheet": stg_sn, "modified": [], "added": []}

    # staging 有、target 没有的列 → 警告并跳过（仅在该列实际有数据时报告）
    for idx, col in enumerate(columns):
        sig = col["sig"]
        if sig in stg_sig_to_phys and sig not in tgt_sig_to_phys:
            has_data = any(
                not _is_empty_cell(*cell)
                for rn in aligned_stg_rows
                for cell in [aligned_stg_rows[rn].get(idx + 1, (None, None))]
            )
            if has_data:
                stg_phys = stg_sig_to_phys[sig]
                meta = stg_phys_to_meta.get(stg_phys, {})
                name = meta.get("row2") or meta.get("row1") or f"第{stg_phys}列"
                report["skipped_columns"].append({
                    "header": name,
                    "reason": "目标表无对应列，该列数据未写入",
                })

    for rn, s_row in aligned_stg_rows.items():
        stg_row_idx = _row_of(stg_sheet_origins, rn)
        t_row = aligned_tgt_rows.get(rn)

        if t_row is None:
            # ---- 新增行：追加到目标 sheet 末尾，只写映射列 ----
            writes = []
            for idx, col in enumerate(columns):
                sig = col["sig"]
                tgt_phys = tgt_sig_to_phys.get(sig)
                if tgt_phys is None:
                    continue  # staging 独有列，跳过
                cell = s_row.get(idx + 1)
                if cell is None:
                    continue
                s_val, s_fml = cell
                if _is_empty_cell(s_val, s_fml):
                    continue
                writes.append({
                    "col": tgt_phys,
                    "stg_col": stg_sig_to_phys.get(sig),
                    "value": s_val,
                    "formula": s_fml,
                })
            if writes:
                report["added"].append(rn)
                plan["added"].append({
                    "rowname": rn,
                    "stg_row": stg_row_idx,
                    "writes": writes,
                })
            continue

        # ---- 已有行：staging 非空单元格覆盖 target ----
        tgt_row_idx = _row_of(tgt_sheet_origins, rn)
        changes = []
        for idx, col in enumerate(columns):
            sig = col["sig"]
            tgt_phys = tgt_sig_to_phys.get(sig)
            if tgt_phys is None:
                continue
            cell = s_row.get(idx + 1)
            if cell is None:
                continue
            s_val, s_fml = cell
            if _is_empty_cell(s_val, s_fml):
                continue  # staging 为空 → 不动 target，防止误删数据
            t_val, t_fml = t_row.get(idx + 1, (None, None))

            if s_fml:
                # 公式按各自真实位置归一化后比较（=B4 在 A4 与 =B14 在 A14 同义）
                stg_phys = stg_sig_to_phys.get(sig, tgt_phys)
                s_sig = normalize_formula_for_compare(
                    s_fml, stg_row_idx or 0, stg_phys
                )
                t_sig = normalize_formula_for_compare(
                    t_fml, tgt_row_idx or 0, tgt_phys
                ) if t_fml else ""
                if s_sig == t_sig:
                    continue
                new_disp = _disp(s_val, s_fml)
            else:
                if _cell_equal(s_val, t_val):
                    continue  # 值相同 → 未变（target 上的公式也保留不覆盖）
                new_disp = _disp(s_val, None)

            stg_phys = stg_sig_to_phys.get(sig)
            meta = stg_phys_to_meta.get(stg_phys, {})
            changes.append({
                "col": tgt_phys,
                "header": meta.get("row2") or meta.get("row1") or f"第{tgt_phys}列",
                "old": _disp(t_val, t_fml),
                "new": new_disp,
                "value": s_val,
                "formula": s_fml,
                "stg_row": stg_row_idx,
            })

        if changes:
            report["modified"].append({
                "rowname": rn,
                "cells": [{k: c[k] for k in ("col", "header", "old", "new")} for c in changes],
            })
            plan["modified"].append({"rowname": rn, "cells": changes})
        else:
            report["unchanged"].append(rn)

    return report, plan


def _build_merge_plan(staging_path, target_path):
    """读取两表并计算合并计划。返回 (report, plans, tgt_origins, extlink_cache)。"""
    stg_data, stg_headers, stg_origins = _read_all_rows_with_formulas(
        str(staging_path), include_origins=True
    )
    tgt_data, tgt_headers, tgt_origins = _read_all_rows_with_formulas(
        str(target_path), include_origins=True
    )

    # 写回前校验：staging 全部 # sheet 不允许重复 RowName
    stg_dups = _find_duplicate_rownames(staging_path)
    if stg_dups:
        details = "; ".join(f"{d['sheet']}!{d['rowname']}x{d['count']}" for d in stg_dups)
        raise AddMergeError(f"暂存表存在重复 RowName，中止合并: {details}")

    # sheet 匹配（大小写不敏感）
    tgt_wb = openpyxl.load_workbook(target_path, read_only=True)
    try:
        tgt_sheetnames = list(tgt_wb.sheetnames)
    finally:
        tgt_wb.close()

    report = {
        "staging": str(staging_path),
        "target": str(target_path),
        "apply": False,
        "sheets": OrderedDict(),
        "errors": [],
    }
    plans = OrderedDict()
    matched_tgt_sheets = []

    for stg_sn in stg_data.keys():
        tgt_sn = _match_target_sheet(stg_sn, tgt_sheetnames)
        if tgt_sn is None:
            report["sheets"][stg_sn] = {
                "staging_sheet": stg_sn,
                "target_sheet": None,
                "status": "skipped",
                "skip_reason": "目标表中没有同名 sheet（不新建 sheet，已保守跳过）",
                "added": [], "modified": [], "unchanged": [], "skipped_columns": [],
            }
            continue
        matched_tgt_sheets.append(tgt_sn)
        sheet_report, sheet_plan = _plan_sheet(
            stg_sn, tgt_sn, stg_data, stg_headers, stg_origins,
            tgt_data, tgt_headers, tgt_origins,
        )
        report["sheets"][stg_sn] = sheet_report
        plans[tgt_sn] = sheet_plan

    # target 被匹配的 sheet 不允许重复 RowName
    tgt_dups = _find_duplicate_rownames(target_path, sheets=set(matched_tgt_sheets))
    if tgt_dups:
        details = "; ".join(f"{d['sheet']}!{d['rowname']}x{d['count']}" for d in tgt_dups)
        raise AddMergeError(f"目标表存在重复 RowName，中止合并: {details}")

    # 外链缓存必须在 openpyxl 写回之前抓取
    extlink_cache = _read_external_link_cache(str(target_path))
    return report, plans, tgt_origins, extlink_cache


# ---------------------------------------------------------------------------
# 写回
# ---------------------------------------------------------------------------

def _last_content_row(ws):
    """目标 sheet 最后一个含任意内容的行号（无内容则返回 DATA_START_ROW-1）。"""
    last = DATA_START_ROW - 1
    for row_idx, row_cells in enumerate(
        ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row),
        start=DATA_START_ROW,
    ):
        for cell in row_cells:
            if cell.value is not None and str(cell.value).strip() != "":
                last = row_idx
                break
    return last


def _read_metadata_xml(filepath):
    """Cache xl/metadata.xml (XLDAPR) before openpyxl drops it on save."""
    if not os.path.isfile(filepath):
        return None
    try:
        with zipfile.ZipFile(filepath) as z:
            if "xl/metadata.xml" in z.namelist():
                data = z.read("xl/metadata.xml")
                if b"XLDAPR" in data:
                    return data
    except Exception:
        return None
    return None


def _inject_cached_metadata(metadata_xml, target_path):
    """Restore xl/metadata.xml from bytes captured before the openpyxl save."""
    if not metadata_xml:
        return
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".zip")
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/metadata.xml", metadata_xml)
        _inject_dynamic_array_metadata(tmp_path, str(target_path))
    finally:
        if os.path.isfile(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def _write_changes(target_path, plans, tgt_origins, extlink_cache, staging_path=None):
    """写回目标表：COM 直写（静态值）优先；无 Excel 降级 openpyxl 写回。

    COM 路径：openpyxl 仅用于计算写入坐标（不保存）→ write_cells_via_com。
      公式单元格（formula 字段）→ AddMergeError 拒绝（COM 直写不支持公式）；
      COM Save 天然保留外链/动态数组元数据，无需 zip 注入。
    降级路径：openpyxl 全量写回 + zip 层注入外链/动态数组元数据，并明确警告
      （公式缓存将被清空、外链需注入恢复、文件可能无法被 Excel 正常打开）。
    返回 {"method": "com" | "openpyxl-fallback", "warning": str | None}。
    """
    target_path = str(target_path)
    metadata_xml = _read_metadata_xml(target_path)
    stg_wb = None
    if staging_path and os.path.isfile(staging_path):
        try:
            stg_wb = openpyxl.load_workbook(staging_path, read_only=True, data_only=False)
        except Exception:
            stg_wb = None
    try:
        wb = openpyxl.load_workbook(target_path)
        try:
            # ---- 只读计算写入计划：modified 目标坐标 / added 插入位置 ----
            # cells[sheet] = [(row, col, value, formula, stg_row, stg_col)]，
            # formula 非空表示公式写入（COM 路径拒绝，降级路径保留公式语义）。
            cells = {}
            for tgt_sn, sheet_plan in plans.items():
                ws = wb[tgt_sn]
                origins = tgt_origins.get(tgt_sn, {})
                stg_sheet = sheet_plan.get("staging_sheet")
                sheet_cells = cells.setdefault(tgt_sn, [])

                def _staging_cell_is_array(stg_row, stg_col):
                    if stg_wb is None or not stg_sheet or stg_sheet not in stg_wb.sheetnames:
                        return False
                    try:
                        return isinstance(
                            stg_wb[stg_sheet].cell(row=stg_row, column=stg_col).value,
                            ArrayFormula,
                        )
                    except Exception:
                        return False

                # 修改：只覆盖 staging 非空的映射列单元格
                for mod in sheet_plan["modified"]:
                    t_row = _row_of(origins, mod["rowname"])
                    if t_row is None:
                        raise AddMergeError(
                            f"内部错误: 目标 sheet {tgt_sn} 找不到行 '{mod['rowname']}' 的坐标"
                        )
                    for ch in mod["cells"]:
                        sheet_cells.append((
                            t_row, ch["col"], ch["value"],
                            ch.get("formula") or "", ch.get("stg_row"), None,
                        ))

                # 新增：追加到最后一个内容行之后，只写映射列
                next_row = _last_content_row(ws) + 1
                for add in sheet_plan["added"]:
                    for w in add["writes"]:
                        sheet_cells.append((
                            next_row, w["col"], w["value"],
                            w.get("formula") or "", add["stg_row"],
                            w.get("stg_col"),
                        ))
                    next_row += 1

            # ---- 写回方式判定 ----
            from excel_recalc import com_available
            if com_available():
                com_writes = {}
                for tgt_sn, sheet_cells in cells.items():
                    for row, col, value, formula, _stg_row, _stg_col in sheet_cells:
                        if formula:
                            raise AddMergeError(
                                f"staging 含公式单元格，COM 直写不支持公式写入，"
                                f"请改为静态值: {tgt_sn} r{row}c{col}"
                            )
                        com_writes.setdefault(tgt_sn, []).append((row, col, value))
                from excel_recalc import write_cells_via_com
                write_cells_via_com(target_path, com_writes)
                return {"method": "com", "warning": None}

            # ---- 降级：openpyxl 全量写回（公式/数组公式保留原语义） ----
            for tgt_sn, sheet_cells in cells.items():
                ws = wb[tgt_sn]
                stg_sheet = plans[tgt_sn].get("staging_sheet")

                def _staging_cell_is_array(stg_row, stg_col):
                    if stg_wb is None or not stg_sheet or stg_sheet not in stg_wb.sheetnames:
                        return False
                    try:
                        return isinstance(
                            stg_wb[stg_sheet].cell(row=stg_row, column=stg_col).value,
                            ArrayFormula,
                        )
                    except Exception:
                        return False

                for row, col, value, formula, stg_row, stg_col in sheet_cells:
                    if formula:
                        # 公式中的相对行号从 staging 行映射到 target 行
                        row_map = {stg_row: row} if stg_row else {}
                        value = adjust_formula(formula, row_map)
                        if isinstance(ws.cell(row=row, column=col).value, ArrayFormula) \
                                or _staging_cell_is_array(stg_row, stg_col):
                            # 目标单元格是数组公式，覆盖后必须保留数组上下文
                            value = ArrayFormula(
                                f"{get_column_letter(col)}{row}", value
                            )
                    ws.cell(row=row, column=col, value=value)

            wb.save(target_path)
        finally:
            try:
                wb.close()
            except Exception:
                pass
    finally:
        if stg_wb is not None:
            try:
                stg_wb.close()
            except Exception:
                pass

    # openpyxl 保存会损坏外链 XML / 丢失动态数组元数据，zip 层恢复
    if extlink_cache and extlink_cache.get("files"):
        _inject_external_links(target_path, target_path, cache=extlink_cache)
    # 目标自身保存后已无 metadata.xml，必须用写回前缓存的原始内容
    _inject_cached_metadata(metadata_xml, target_path)

    warning = (
        "本机无 win32com/Excel，已降级 openpyxl 写回：公式缓存将被清空、"
        "外部链接需注入恢复、文件可能无法被 Excel 正常打开。"
        "建议在装有 Excel 的机器上重新执行合并，或人工用交互式 Excel "
        "打开→点「更新」→保存修复缓存。"
    )
    print(f"[add-merge] 降级警告: {warning}", file=sys.stderr)
    return {"method": "openpyxl-fallback", "warning": warning}


def _verify_external_links_preserved(extlink_cache, target_path):
    """写回前存在的外链，写回后必须仍在，否则视为损坏。"""
    before = (extlink_cache or {}).get("files", {}) or {}
    if not before:
        return
    after_cache = _read_external_link_cache(str(target_path)) or {}
    after = after_cache.get("files", {}) or {}
    missing = sorted(name for name in before if name not in after)
    if missing:
        raise AddMergeError(f"外部链接丢失: {', '.join(missing[:5])}")



# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def apply_additive_merge(staging_path, target_path, apply=False, recalc_timeout=120):
    """增量合并：staging → target（只增只改，绝不删行）。

    apply=False（默认）只生成报告不写文件；apply=True 走事务写回。
    返回报告 dict：
      sheets: {staging_sheet: {target_sheet, status, added, modified,
                               unchanged, skipped_columns, skip_reason}}
      written / backup_path / na_guard（apply 时）
    """
    staging_path = Path(staging_path)
    target_path = Path(target_path)
    if not staging_path.exists():
        raise AddMergeError(f"暂存表不存在: {staging_path}")
    if not target_path.exists():
        raise AddMergeError(f"目标表不存在: {target_path}")

    report, plans, tgt_origins, extlink_cache = _build_merge_plan(
        staging_path, target_path
    )
    report["apply"] = bool(apply)

    total_changes = sum(
        len(s["added"]) + len(s["modified"]) for s in report["sheets"].values()
    )
    if not apply or total_changes == 0:
        report["written"] = False
        return report

    # ---- 事务写回：备份 → 写回前基线 → 写回 → 外链校验 → #N/A 基线校验 ----
    backup_path = create_merge_backup(target_path)
    try:
        from excel_recalc import count_formula_error_caches, verify_formula_cache_guard
        baseline = count_formula_error_caches(str(target_path))
        write_info = _write_changes(
            target_path, plans, tgt_origins, extlink_cache, staging_path=staging_path
        )
        _verify_external_links_preserved(extlink_cache, target_path)
        if write_info.get("method") == "com":
            na_guard = verify_formula_cache_guard(str(target_path), baseline)
        else:
            # 降级路径：openpyxl 写回已清空缓存，错误值统计无意义
            na_guard = {
                "ok": None,
                "skipped": write_info.get("warning") or "已降级 openpyxl 写回，跳过 #N/A 基线校验",
            }
    except Exception as exc:
        restored = False
        if backup_path:
            restored = restore_merge_backup(backup_path, target_path)
        msg = f"增量合并写回失败: {exc}"
        msg += "；已从备份恢复目标表" if restored else "；备份恢复失败，请手工处理"
        raise AddMergeError(msg, backup_path=backup_path, restored=restored) from exc

    report["written"] = True
    report["backup_path"] = str(backup_path) if backup_path else None
    report["na_guard"] = na_guard
    return report


# ---------------------------------------------------------------------------
# 报告格式化（中文输出）
# ---------------------------------------------------------------------------

def format_report(report):
    """把 apply_additive_merge 的报告格式化为中文文本。"""
    lines = []
    mode = "APPLY（已写入）" if report.get("written") else "DRY-RUN（未写入）"
    lines.append("=" * 70)
    lines.append(f"  增量合并报告 — {mode}")
    lines.append("=" * 70)
    lines.append(f"  暂存表: {report['staging']}")
    lines.append(f"  目标表: {report['target']}")

    total_added = total_modified = total_unchanged = 0
    for stg_sn, s in report["sheets"].items():
        lines.append(f"\n  [SHEET] {stg_sn}")
        if s["status"] == "skipped":
            lines.append(f"    [跳过] {s['skip_reason']}")
            continue
        if s["target_sheet"] != stg_sn:
            lines.append(f"    → 匹配目标 sheet: {s['target_sheet']}")

        added = s["added"]
        modified = s["modified"]
        unchanged = s["unchanged"]
        total_added += len(added)
        total_modified += len(modified)
        total_unchanged += len(unchanged)

        lines.append(f"    新增 {len(added)} 行 / 修改 {len(modified)} 行 / 未变 {len(unchanged)} 行")

        for rn in added:
            lines.append(f"      [新增] {rn}")
        for mod in modified:
            lines.append(f"      [修改] {mod['rowname']} ({len(mod['cells'])} 个单元格)")
            for c in mod["cells"]:
                lines.append(
                    f"          {c['header']}: {c['old'] or '(空)'} → {c['new'] or '(空)'}"
                )
        if unchanged:
            preview = ", ".join(str(rn) for rn in unchanged[:10])
            more = f" 等 {len(unchanged)} 行" if len(unchanged) > 10 else ""
            lines.append(f"      [未变] {preview}{more}")
        for sc in s["skipped_columns"]:
            lines.append(f"      [警告] 列 '{sc['header']}': {sc['reason']}")

    lines.append(f"\n{'─' * 70}")
    lines.append(
        f"  总计: 新增 {total_added} 行, 修改 {total_modified} 行, 未变 {total_unchanged} 行"
    )
    if report.get("written"):
        lines.append(f"  [OK] 已写入目标表。备份: {report.get('backup_path')}")
        na_guard = report.get("na_guard") or {}
        if na_guard.get("skipped"):
            lines.append(f"  [WARN] {na_guard['skipped']}")
        elif na_guard.get("ok") is False:
            sheets = ", ".join(na_guard.get("added_sheets") or {})
            lines.append(
                f"  [WARN] 写回后新增 {na_guard.get('total')} 个错误值"
                + (f"（{sheets}）" if sheets else "")
                + "；请策划检查对应表数据"
            )
        elif na_guard.get("ok") is True:
            lines.append("  [OK] 写回后无新增错误值（#N/A/#REF! 基线校验通过）。")
    else:
        if total_added or total_modified:
            lines.append("  未写入任何文件；确认无误后加 --apply 执行。")
        else:
            lines.append("  没有任何变更需要写入。")
    return "\n".join(lines)
