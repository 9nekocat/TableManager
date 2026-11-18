# -*- coding: utf-8 -*-
"""
妙妙工具箱后端逻辑（TableManager 分区功能）

工具 1：Effect #N/A 修复重算
  - 针对 Effect.xlsx 的 Params/ConditionIDs 公式列（TEXTJOIN+XLOOKUP）
  - 把 IF(X="无",0, 升级为 IF(OR(X="无",X=""),0,（覆盖空单元格 → XLOOKUP(0) #N/A 坑）
  - COM 读回写 + CalculateFull + Excel 原生保存
  - 前置：备份（Raw 外）、占用检查；后置：错误缓存校验 + 残留明细

工具 2：查 #N/A（通用）
  - 对任意 Raw xlsx 统计公式错误缓存（#N/A/#REF!/#VALUE! 等）并列出明细
  - XML 层解析（同 excel_recalc.count_formula_error_caches），不打开 Excel
"""
import os
import shutil
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import excel_recalc  # noqa: E402

try:
    import effect_recalc  # noqa: E402
except Exception:  # pragma: no cover
    effect_recalc = None

DEFAULT_EFFECT = r"A:/DataTable/Raw/Effect.xlsx"
BACKUP_DIR = r"A:/DataTable/Raw/../effect_recalc_backups"


# ===== 工具 1：Effect #N/A 修复重算 =====

def effect_status(path=None):
    """Effect.xlsx 当前错误缓存统计 + 占用状态（不修改任何文件）"""
    path = path or DEFAULT_EFFECT
    result = {
        "ok": True,
        "path": path,
        "exists": os.path.isfile(path),
        "open_in_excel": _is_open_in_excel(path),
        "error_counts": {},
        "total_errors": 0,
        "residual_details": [],
    }
    if not result["exists"]:
        result["ok"] = False
        result["error"] = "文件不存在: %s" % path
        return result
    counts = excel_recalc.count_formula_error_caches(path)
    result["error_counts"] = counts
    result["total_errors"] = sum(counts.values())
    if result["total_errors"]:
        result["residual_details"] = effect_recalc.extract_error_details(path)[:500]
    return result


def effect_recalc_run(path=None, backup=True):
    """执行 Effect 修复重算：备份 → COM 读回写 + OR 保护 → CalculateFull → 保存 → 校验"""
    path = path or DEFAULT_EFFECT
    if not os.path.isfile(path):
        return {"ok": False, "error": "文件不存在: %s" % path}
    if _is_open_in_excel(path):
        return {
            "ok": False,
            "error": "文件正被 Excel 打开（存在 ~$ 锁文件），请先关闭再执行",
        }

    before = excel_recalc.count_formula_error_caches(path)
    backup_path = None
    if backup:
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(BACKUP_DIR, "Effect_%s.xlsx" % ts)
            shutil.copy2(path, backup_path)
        except Exception as e:
            return {"ok": False, "error": "备份失败: %s" % e}

    try:
        result = effect_recalc.recalc(path)
    except Exception as e:
        return {
            "ok": False,
            "error": "COM 重算失败: %s" % e,
            "backup_path": backup_path,
            "before": before,
        }

    after = excel_recalc.count_formula_error_caches(path)
    residual = effect_recalc.extract_error_details(path)[:500] if sum(after.values()) else []
    return {
        "ok": True,
        "before": before,
        "after": after,
        "or_protections": result.get("or_protections", 0),
        "backup_path": backup_path,
        "residual_count": len(residual),
        "residual_details": residual,
        "note": (
            "残留 #N/A 为数据缺失类（名称与目标表不一致，如 '寒冷名刀' 不在 "
            "Localization/Buff.xlsx D 列），需改名称列或补目标表数据。"
            if residual else "重算完成，无残留 #N/A。"
        ),
    }


# ===== 工具 2：查 #N/A（通用，任意 Raw 表） =====

def check_na(path):
    """对任意 xlsx 统计并列出公式错误缓存明细"""
    if not os.path.isfile(path):
        return {"ok": False, "error": "文件不存在: %s" % path}
    try:
        counts = excel_recalc.count_formula_error_caches(path)
    except Exception as e:
        return {"ok": False, "error": "解析失败: %s" % e}
    details = []
    if sum(counts.values()):
        try:
            details = effect_recalc.extract_error_details(path)[:1000]
        except Exception:
            details = []
    return {
        "ok": True,
        "path": path,
        "error_counts": counts,
        "total_errors": sum(counts.values()),
        "details": details,
    }


def _is_open_in_excel(path):
    """检查文件是否正被 Excel 打开（~$ 锁文件）"""
    d = os.path.dirname(path)
    name = os.path.basename(path)
    lock = os.path.join(d, "~$" + name)
    return os.path.exists(lock)
