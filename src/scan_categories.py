# -*- coding: utf-8 -*-
"""
categories.json 校准工具。

实时扫描 RAW_DIR 下的实际 xlsx 文件，与 categories.json 对照：
  - 配置中存在但磁盘缺失的文件（会导致引用检查悬空、校验警告）
  - 磁盘上存在但未编目的文件（不会被校验/合并/引用扫描覆盖）

用法:
  python scan_categories.py           只报告，不修改
  python scan_categories.py --write   写回 categories.json：删除缺失条目，
                                      未编目文件加入 uncategorized（先自动备份）
"""

import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

# 源码运行时把 src/ 自身加入模块搜索路径（打包后由 PyInstaller 处理）
sys.path.insert(0, str(Path(__file__).resolve().parent))

import manager_core as mc

# Excel 临时锁文件、临时目录、隐藏目录一律跳过
SKIP_FILE_PREFIXES = ("~$",)


def scan_raw_files():
    """RAW_DIR 下全部 xlsx 的相对路径（posix 风格）。"""
    files = []
    for p in sorted(mc.RAW_DIR.rglob("*.xlsx")):
        if p.name.startswith(SKIP_FILE_PREFIXES):
            continue
        rel = p.relative_to(mc.RAW_DIR).as_posix()
        if any(part.startswith(".") for part in rel.split("/")):
            continue
        files.append(rel)
    return files


def load_raw_categories():
    """读取原始 categories.json（保留 _comment 等字段与结构）。"""
    try:
        return json.loads(mc.CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[ERROR] 无法读取 {mc.CONFIG_PATH}: {e}")
        sys.exit(1)


def diff_categories(raw_cfg, disk_files):
    """返回 (missing, uncataloged)：配置有而磁盘无 / 磁盘有而配置无。"""
    config_files = []
    for cat in raw_cfg.get("categories", {}).values():
        config_files.extend(cat.get("files", []))
    if raw_cfg.get("uncategorized"):
        config_files.extend(raw_cfg["uncategorized"].get("files", []))
    config_set = set(config_files)
    disk_set = set(disk_files)
    missing = [f for f in config_files if f not in disk_set]
    uncataloged = [f for f in disk_files if f not in config_set]
    return missing, uncataloged


def apply_write(raw_cfg, missing, uncataloged):
    """删缺失条目、未编目文件加入 uncategorized，备份后写回。"""
    missing_set = set(missing)
    for key, cat in raw_cfg.get("categories", {}).items():
        kept = [f for f in cat.get("files", []) if f not in missing_set]
        cat["files"] = kept
    if uncataloged:
        uncat = raw_cfg.setdefault("uncategorized", {
            "display_name": "未分类",
            "description": "尚未归入分类的文件",
            "files": [],
        })
        uncat.setdefault("files", [])
        for f in uncataloged:
            if f not in uncat["files"]:
                uncat["files"].append(f)

    backup = mc.CONFIG_PATH.with_suffix(
        f".bak-{datetime.now():%Y%m%d%H%M%S}.json"
    )
    shutil.copy2(mc.CONFIG_PATH, backup)
    mc.CONFIG_PATH.write_text(
        json.dumps(raw_cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return backup


def main():
    write = "--write" in sys.argv
    disk_files = scan_raw_files()
    raw_cfg = load_raw_categories()
    missing, uncataloged = diff_categories(raw_cfg, disk_files)

    print("=" * 70)
    print("  categories.json 校准 — 扫描 DataTable/Raw")
    print("=" * 70)
    print(f"  磁盘 xlsx: {len(disk_files)} 个")

    if missing:
        print(f"\n  [配置中存在但磁盘缺失] ({len(missing)})")
        for f in missing:
            print(f"    - {f}")
    else:
        print("\n  [OK] 没有失效的配置条目")

    if uncataloged:
        print(f"\n  [磁盘存在但未编目] ({len(uncataloged)})")
        for f in uncataloged:
            print(f"    - {f}")
    else:
        print("  [OK] 没有漏扫描的文件")

    if write:
        if not missing and not uncataloged:
            print("\n  无需修改。")
            return
        backup = apply_write(raw_cfg, missing, uncataloged)
        print(f"\n  [OK] categories.json 已更新（备份: {backup.name}）")
        if uncataloged:
            print("  新文件已加入 uncategorized，建议人工归入合适分类。")
    elif missing or uncataloged:
        print("\n  加 --write 可将上述结果写回 categories.json。")

    if missing:
        sys.exit(1)


if __name__ == "__main__":
    main()
