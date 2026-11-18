# -*- coding: utf-8 -*-
"""
DataTable Manager CLI
管理 DataTable/Raw 下的表格：分类浏览、快速打开、错误值校验
SVN 集成、冲突检测、合并

用法:
    python table_manager.py list             列出所有表格（分类显示）
    python table_manager.py list -c Cards    按类别筛选
    python table_manager.py open Card        在 Excel 中打开表格
    python table_manager.py info Effect      查看表格详情
    python table_manager.py check            导表检查（错误值/重复ID/表头/类型/引用）
    python table_manager.py check -f Effect  校验单个文件
    python table_manager.py check --csv out  导出 CSV 报告
    python table_manager.py refs Effect      查看谁引用了 Effect.xlsx
    python table_manager.py refs --dangling  列出悬空引用
    python table_manager.py svn status       SVN 状态
    python table_manager.py svn update       SVN 更新
    python table_manager.py svn log          SVN 日志
    python table_manager.py merge status     列出冲突文件
    python table_manager.py merge preview <f> 生成合并预览
    python table_manager.py merge resolve <f> 应用合并结果
    python table_manager.py merge abort <f>   放弃合并
"""

import argparse
import os
import sys
from pathlib import Path

# 源码运行时把 src/ 自身加入模块搜索路径（打包后由 PyInstaller 处理）
sys.path.insert(0, str(Path(__file__).resolve().parent))

from manager_core import (
    RAW_DIR,
    TEMP_DIR,
    CONFIG_PATH,
    load_categories,
    discover_tables,
    find_table,
    open_in_excel,
    get_table_info,
    validate_all,
    format_table_list,
    format_validation_report,
    format_table_info,
)
from svn_client import SvnClient, SvnError
from merge_engine import generate_merge_preview
from addmerge_engine import AddMergeError, apply_additive_merge, format_report
from merge_workflow import (
    MergeWorkflowError,
    apply_merge_preview_transactional,
    recover_merge_preview_transactional,
    find_merge_preview,
    original_name_from_preview,
    infer_theirs_path,
)


# ====== list 子命令 ======
def cmd_list(args):
    config = load_categories()
    by_cat, uncataloged = discover_tables(config)
    output = format_table_list(by_cat, config, category_filter=args.category)
    print(output)
    if uncataloged:
        print(f"\n[WARN] 以下文件未在 categories.json 中配置:")
        for f in uncataloged:
            print(f"    - {f}")
        print(f"  共 {len(uncataloged)} 个未编目文件。")


# ====== open 子命令 ======
def cmd_open(args):
    config = load_categories()
    fp, info = find_table(args.name, config)
    if not fp:
        direct = RAW_DIR / args.name
        if direct.exists():
            fp = args.name
        else:
            print(f"[ERROR] 未找到匹配的表格: {args.name}")
            return
    print(f"[>] 打开: {RAW_DIR / fp}")
    open_in_excel(fp)


# ====== info 子命令 ======
def cmd_info(args):
    config = load_categories()
    fp, info = find_table(args.name, config)
    if not fp:
        direct = RAW_DIR / args.name
        if direct.exists():
            fp = args.name
            info = get_table_info(fp)
            info["category"] = config.get("by_file", {}).get(fp, "未分类")
        else:
            print(f"[ERROR] 未找到匹配的表格: {args.name}")
            return
    print(format_table_info(info))


# ====== check 子命令 ======
def cmd_check(args):
    config = load_categories()
    print("[*] 开始校验...")
    print(f"   扫描路径: {RAW_DIR}")
    if args.file:
        print(f"   文件过滤: {args.file}")
    errors = validate_all(config, file_filter=args.file)
    csv_out = args.csv
    if csv_out and not csv_out.endswith(".csv"):
        csv_out += ".csv"
    report = format_validation_report(errors, csv_file=csv_out)
    print(report)
    if errors:
        sys.exit(1)


# ====== refs 子命令 ======
def cmd_refs(args):
    import ref_index

    if not ref_index.is_available():
        print("[ERROR] 未找到 references.json 或其中没有规则，引用查找不可用。")
        return

    print("[*] 构建引用索引（首次较慢）...")
    index = ref_index.build_reference_index()
    print(f"    共 {len(index.records)} 条引用，{len(index.dangling)} 条悬空。")

    if args.dangling:
        records = index.dangling
        if args.target:
            config = load_categories()
            fp, _ = find_table(args.target, config)
            if not fp:
                print(f"[ERROR] 未找到匹配的表格: {args.target}")
                return
            records = [r for r in records if r["src_file"] == fp]
        print("=" * 80)
        print(f"  悬空引用 — {len(records)} 条")
        print("=" * 80)
        for r in records:
            print(
                f"  {r['src_file']} {r['src_sheet']}!{r['src_col_ref']} "
                f"[{r['rule']}] '{r['value']}' 目标不存在 ({r['level']})"
            )
        if records:
            sys.exit(1)
        return

    if not args.target:
        print("[ERROR] 请指定目标表（或使用 --dangling 查看悬空引用）。")
        return

    config = load_categories()
    fp, _ = find_table(args.target, config)
    if not fp:
        direct = RAW_DIR / args.target
        if direct.exists():
            fp = args.target
        else:
            print(f"[ERROR] 未找到匹配的表格: {args.target}")
            return

    records = index.find_references(fp, sheet=args.sheet, row=args.row)
    print("=" * 80)
    print(f"  引用查询 — {fp}" + (f" {args.sheet}" if args.sheet else "") + (f" 行{args.row}" if args.row else ""))
    print(f"  共 {len(records)} 个引用来源")
    print("=" * 80)
    # 按来源文件分组
    by_src = {}
    for r in records:
        by_src.setdefault(r["src_file"], []).append(r)
    for src, recs in sorted(by_src.items()):
        print(f"\n  [{src}] ({len(recs)})")
        for r in recs[: args.limit]:
            targets = ", ".join(f"{t[1]}!A{t[2]}" for t in r["targets_hit"])
            print(
                f"    {r['src_sheet']}!{r['src_col_ref']} [{r['kind']}/{r['rule']}] "
                f"'{r['value']}' -> {targets}"
            )
        if len(recs) > args.limit:
            print(f"    ... 还有 {len(recs) - args.limit} 条（--limit 调整）")


# ====== pcheck 子命令 ======
def cmd_pcheck(args):
    import processed_checks

    if not processed_checks.is_available():
        print("[ERROR] 未找到 processed_checks.json 或 DataTable/Processed 目录不存在。")
        return

    result = processed_checks.run_checks(table_filter=args.table)
    stats = result["stats"]
    print("=" * 80)
    print("  导出数据检查 — DataTable/Processed (DT_*.json)")
    print(f"  共 {stats['total_records']} 条引用，{stats['total_dangling']} 条悬空")
    print("=" * 80)

    # 按来源表分组
    by_src = {}
    for r in result["dangling"]:
        by_src.setdefault(r["src_file"], []).append(r)
    for src in sorted(by_src):
        recs = by_src[src]
        print(f"\n  [{src}] ({len(recs)})")
        for r in recs[: args.limit]:
            targets = ", ".join(t[0] for t in r["targets_hit"])
            level_tag = " (warning)" if r["level"] == "warning" else ""
            print(
                f"    行{r['src_row']} {r['src_col_ref']} [{r['rule']}] "
                f"'{r['value']}' 目标不存在{level_tag}"
            )
        if len(recs) > args.limit:
            print(f"    ... 还有 {len(recs) - args.limit} 条（--limit 调整）")

    if result["skipped"]:
        print("\n  [跳过的规则]")
        for s in result["skipped"]:
            print(f"    - {s['rule']}: {s['reason']}")

    if args.csv:
        import csv as csv_mod
        out = args.csv if args.csv.endswith(".csv") else args.csv + ".csv"
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv_mod.DictWriter(
                f, fieldnames=["src_file", "src_row", "src_col_ref", "rule", "value", "level"]
            )
            writer.writeheader()
            writer.writerows(result["dangling"])
        print(f"\n[CSV] 悬空引用已导出至: {out}")

    if any(r["level"] == "error" for r in result["dangling"]):
        sys.exit(1)


# ====== svn status ======
def cmd_svn_status(args):
    svn = SvnClient()
    try:
        is_wc = svn.is_working_copy()
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    if not is_wc:
        print("[ERROR] DataTable/Raw 不是 SVN 工作副本。")
        return

    try:
        items = svn.status()
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    if not items:
        print("[OK] 没有变更，工作副本干净。")
        return

    modified = [i for i in items if i["status"] == "M"]
    conflicted = [i for i in items if i["conflicted"]]
    added = [i for i in items if i["status"] == "A"]
    deleted = [i for i in items if i["status"] == "D"]
    unversioned = [i for i in items if i["status"] == "?"]

    print("=" * 70)
    print("  SVN 状态 — DataTable/Raw")
    print("=" * 70)

    if conflicted:
        print(f"\n  [冲突] ({len(conflicted)})")
        for item in conflicted:
            print(f"    C  {item['path']}")

    if modified:
        print(f"\n  [已修改] ({len(modified)})")
        for item in modified:
            print(f"    M  {item['path']}")

    if added:
        print(f"\n  [新增] ({len(added)})")
        for item in added:
            print(f"    A  {item['path']}")

    if deleted:
        print(f"\n  [已删除] ({len(deleted)})")
        for item in deleted:
            print(f"    D  {item['path']}")

    if unversioned:
        print(f"\n  [未版本控制] ({len(unversioned)})")
        for item in unversioned:
            print(f"    ?  {item['path']}")

    print(f"\n  总计: {len(items)} 个变更")


# ====== svn update ======
def cmd_svn_update(args):
    svn = SvnClient()
    try:
        is_wc = svn.is_working_copy()
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    if not is_wc:
        print("[ERROR] DataTable/Raw 不是 SVN 工作副本。")
        return

    print("[*] 正在 svn update DataTable/Raw ...")
    try:
        summary = svn.update()
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    print(f"  版本: {summary['revision']}")
    print(f"  更新: {summary['updated']}  合并: {summary['merged']}  "
          f"冲突: {summary['conflicted']}  新增: {summary['added']}  删除: {summary['deleted']}")

    if summary["conflicted"] > 0:
        print("\n[WARN] 存在冲突文件！使用 'merge status' 查看，'merge preview <file>' 开始合并。")


# ====== svn log ======
def cmd_svn_log(args):
    svn = SvnClient()
    try:
        is_wc = svn.is_working_copy()
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    if not is_wc:
        print("[ERROR] DataTable/Raw 不是 SVN 工作副本。")
        return

    limit = args.n or 10
    try:
        entries = svn.log(limit=limit)
    except SvnError as e:
        print(f"[ERROR] {e}")
        return

    print("=" * 70)
    print(f"  SVN 日志 — DataTable/Raw (最近 {limit} 条)")
    print("=" * 70)
    for e in entries:
        print(f"\n  {e['revision']} | {e['author']} | {e['date']}")
        print(f"    {e['message']}")


# ====== merge status ======
def cmd_merge_status(args):
    svn = SvnClient()
    conflicts = svn.find_conflict_files()
    if not conflicts:
        print("[OK] 没有冲突文件。")
        return

    print("=" * 70)
    print(f"  冲突文件 — ({len(conflicts)})")
    print("=" * 70)

    for base_rel, entry in conflicts.items():
        print(f"\n  [FILE] {base_rel}")
        if "mine" in entry:
            print(f"    我的版本:   {entry['mine']}")
        if "theirs" in entry:
            print(f"    仓库版本:   {entry['theirs']}")
        if "base" in entry:
            print(f"    基础版本:   {entry['base']}")
        print(f"    操作: merge preview {base_rel}")


# ====== merge preview ======
def cmd_merge_preview(args):
    svn = SvnClient()
    conflicts = svn.find_conflict_files()
    query = args.file

    # Find matching conflict
    target = None
    for base_rel, entry in conflicts.items():
        if query.lower() in base_rel.lower():
            target = (base_rel, entry)
            break

    if not target:
        print(f"[ERROR] 找不到合并预览文件: {query}")
        print(f"  使用 'merge status' 查看所有冲突文件。")
        return

    base_rel, entry = target
    mine_path = entry.get("mine")
    theirs_path = entry.get("theirs")
    base_path = entry.get("base")

    if not mine_path or not theirs_path:
        print(f"[ERROR] 缺少必要的文件版本。")
        print(f"  mine: {mine_path}")
        print(f"  theirs: {theirs_path}")
        return

    # If no base, use mine as base
    if not base_path:
        base_path = mine_path

    output_path = TEMP_DIR / base_rel.replace(".xlsx", "_merge.xlsx")
    os.makedirs(output_path.parent, exist_ok=True)

    print(f"[*] 生成合并预览...")
    print(f"  基础版本: {base_path}")
    print(f"  我的版本: {mine_path}")
    print(f"  仓库版本: {theirs_path}")
    print(f"  输出文件: {output_path}")
    print(f"  冲突默认: {args.conflict_default}")

    try:
        result = generate_merge_preview(
            base_path, mine_path, theirs_path, str(output_path),
            conflict_default=args.conflict_default,
        )
        print(f"\n[OK] 合并预览已生成: {result}")
        print(f"  [我的] 行 = 黄色, [仓库] 行 = 黄色, [最终] 行 = 绿色")
        print(f"  冲突单元格 = 红色高亮")
        print(f"  请在 Excel 中编辑 [最终] 行的值，保存后执行 merge resolve")

        if not args.no_open:
            os.startfile(str(output_path))
    except Exception as e:
        print(f"[ERROR] 生成合并预览失败: {e}")
        import traceback
        traceback.print_exc()


# ====== merge resolve ======
def cmd_merge_resolve(args):
    svn = SvnClient()
    query = args.file

    preview_path = find_merge_preview(query)

    if not preview_path or not preview_path.exists():
        # Try searching in TEMP_DIR
        merge_files = list(TEMP_DIR.glob("*_merge.xlsx")) if TEMP_DIR.exists() else []
        matching = [f for f in merge_files if query.lower() in f.stem.lower()]
        if matching:
            preview_path = matching[0]
        else:
            print(f"[ERROR] 找不到合并预览文件: {preview_path}")
            print(f"  使用 'merge preview {query}' 先创建合并预览。")
            return

    # Original file is in RAW_DIR
    original_name = original_name_from_preview(preview_path)
    output_path = RAW_DIR / original_name

    print(f"[*] 应用合并结果...")
    print(f"  预览文件: {preview_path}")
    print(f"  目标文件: {output_path}")

    # Auto-detect theirs_path from SVN conflict files
    theirs_path = infer_theirs_path(original_name)

    try:
        result = apply_merge_preview_transactional(
            preview_path,
            output_path,
            mine_path=output_path,
            theirs_path=theirs_path,
            mark_svn_resolved=True,
            cleanup=True,
        )
        print(f"\n[OK] 合并完成: {result}")

        try:
            na_guard = result.get("na_guard") or {}
            if na_guard.get("skipped"):
                print(f"  [WARN] {na_guard['skipped']}")
            elif na_guard.get("ok") is False:
                sheets = ", ".join(na_guard.get("added_sheets") or {})
                print(
                    f"  [WARN] 写回后新增 {na_guard.get('total')} 个错误值"
                    + (f"（{sheets}）" if sheets else "")
                    + "，请策划检查对应表数据"
                )
            else:
                print("  写回后无新增错误值（#N/A/#REF! 基线校验通过）。")
        except Exception as e:
            print(f"  [WARN] #N/A 基线校验失败，请用 Excel 打开并保存后再比对: {e}")

        rel_path = str(output_path.relative_to(RAW_DIR))

        # Run svn resolved
        try:
            if not result.get("svn_resolved"):
                raise SvnError("svn resolved was not run")
            print(f"  SVN 冲突已标记为已解决。")
        except SvnError as e:
            print(f"  [WARN] svn resolved 失败: {e}")

        # Clean up conflict artifacts (shared function)
        cleanup_result = result.get("cleanup") or {"deleted": [], "failed": []}
        for fname in cleanup_result.get("deleted", []):
            print(f"  已清理: {fname}")
        for fname in cleanup_result.get("failed", []):
            print(f"  [WARN] 清理失败: {fname}")

    except Exception as e:
        print(f"[ERROR] 合并失败: {e}")
        import traceback
        traceback.print_exc()


# ====== merge recover ======
def _cmd_merge_recover_impl(args, preview_path, output_path, theirs_path):
    print("[*] Recovering changes from merge preview...")
    print(f"  preview: {preview_path}")
    print(f"  target: {output_path}")

    try:
        result = recover_merge_preview_transactional(
            preview_path,
            output_path,
            backup_path=args.backup,
            theirs_path=theirs_path,
        )
        print(f"\n[OK] recovered: {result['output']}")
        print(f"  pre-write backup: {result.get('backup_path') or '(none)'}")
        print(f"  recovery base: {result.get('recovery_base')}")
    except MergeWorkflowError as e:
        print(f"[ERROR] recover failed: {e}")
        if e.backup_path:
            print(f"  backup: {e.backup_path}")
        print(f"  restored original: {'yes' if e.restored else 'no'}")
        if e.restore_error:
            print(f"  restore backup failed: {e.restore_error}")
    except Exception as e:
        print(f"[ERROR] recover failed: {e}")


def cmd_merge_recover(args):
    query = args.file
    preview_path = find_merge_preview(query)
    if not preview_path or not preview_path.exists():
        print(f"[ERROR] 找不到合并预览文件: {query}")
        return

    original_name = original_name_from_preview(preview_path)
    output_path = RAW_DIR / original_name
    theirs_path = infer_theirs_path(original_name)
    return _cmd_merge_recover_impl(args, preview_path, output_path, theirs_path)


# ====== merge abort ======
def cmd_merge_abort(args):
    query = args.file
    preview_path = None

    if query.endswith("_merge.xlsx"):
        preview_path = TEMP_DIR / query
    else:
        preview_path = TEMP_DIR / query.replace(".xlsx", "_merge.xlsx")

    if not preview_path.exists():
        merge_files = list(TEMP_DIR.glob("*_merge.xlsx")) if TEMP_DIR.exists() else []
        matching = [f for f in merge_files if query.lower() in f.stem.lower()]
        if matching:
            preview_path = matching[0]
        else:
            print(f"[ERROR] 找不到合并预览文件。")
            return

    try:
        os.remove(preview_path)
        print(f"[OK] 已删除合并预览: {preview_path.name}")
    except OSError as e:
        print(f"[ERROR] 删除失败: {e}")


# ====== add-merge 子命令 ======
def cmd_add_merge(args):
    staging = Path(args.staging)
    if not staging.exists():
        print(f"[ERROR] 暂存表不存在: {staging}")
        sys.exit(1)

    # 目标表：优先按直接路径，其次 find_table 模糊匹配 DataTable/Raw
    target_arg = args.target
    target_path = None
    direct = Path(target_arg)
    if direct.exists():
        target_path = direct
    else:
        config = load_categories()
        fp, _ = find_table(target_arg, config)
        if fp:
            target_path = RAW_DIR / fp
        elif (RAW_DIR / target_arg).exists():
            target_path = RAW_DIR / target_arg
    if target_path is None:
        print(f"[ERROR] 未找到匹配的目标表: {target_arg}")
        sys.exit(1)

    print(f"[*] 增量合并: {staging} → {target_path}")
    try:
        report = apply_additive_merge(str(staging), str(target_path), apply=args.apply)
    except AddMergeError as e:
        print(f"[ERROR] {e}")
        if e.backup_path:
            print(f"  备份: {e.backup_path}")
        sys.exit(1)
    print(format_report(report))


# ====== 主入口 ======
def main():
    parser = argparse.ArgumentParser(
        description="DataTable Manager - 配置表管理工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python table_manager.py list              列出所有表格
  python table_manager.py list -c Combat    只看战斗类表格
  python table_manager.py open Effect       打开 Effect.xlsx
  python table_manager.py info CardTable    查看 CardTable 详情
  python table_manager.py check             校验所有表格
  python table_manager.py check -f Effect --csv report  校验并导出 CSV
  python table_manager.py refs Effect       查看谁引用了 Effect.xlsx
  python table_manager.py refs --dangling   列出悬空引用
  python table_manager.py pcheck            导出数据检查（Processed）
  python table_manager.py svn status        SVN 状态
  python table_manager.py svn update        更新 DataTable/Raw
  python table_manager.py merge status      查看冲突文件
  python table_manager.py merge preview Effect  生成合并预览
  python table_manager.py merge resolve Effect  应用合并
  python table_manager.py merge abort Effect    放弃合并
  python table_manager.py add-merge FlagConfigRow --staging staging.xlsx          增量合并预览
  python table_manager.py add-merge FlagConfigRow --staging staging.xlsx --apply  实际写入
        """,
    )

    sub = parser.add_subparsers(dest="command", help="子命令")

    # --- list ---
    p_list = sub.add_parser("list", help="列出所有表格（按类别分组）")
    p_list.add_argument("-c", "--category", help="按类别筛选")
    p_list.set_defaults(func=cmd_list)

    # --- open ---
    p_open = sub.add_parser("open", help="在 Excel 中打开表格")
    p_open.add_argument("name", help="文件名（部分匹配）")
    p_open.set_defaults(func=cmd_open)

    # --- info ---
    p_info = sub.add_parser("info", help="查看表格详情")
    p_info.add_argument("name", help="文件名（部分匹配）")
    p_info.set_defaults(func=cmd_info)

    # --- check ---
    p_check = sub.add_parser("check", help="导表检查（错误值/重复ID/表头/类型/引用完整性）")
    p_check.add_argument("-f", "--file", help="仅校验指定文件（部分匹配）")
    p_check.add_argument("--csv", help="导出 CSV 报告路径")
    p_check.set_defaults(func=cmd_check)

    # --- refs ---
    p_refs = sub.add_parser("refs", help="引用关系查找（谁引用了指定表/行）")
    p_refs.add_argument("target", nargs="?", help="目标表文件名（部分匹配）")
    p_refs.add_argument("-s", "--sheet", help="限定目标 sheet（如 #Main）")
    p_refs.add_argument("-r", "--row", type=int, help="限定目标行号")
    p_refs.add_argument("--dangling", action="store_true", help="列出悬空引用（可配合 target 限定来源表）")
    p_refs.add_argument("--limit", type=int, default=50, help="每个来源表最多显示条数 (默认50)")
    p_refs.set_defaults(func=cmd_refs)

    # --- pcheck ---
    p_pcheck = sub.add_parser("pcheck", help="导出数据检查（Processed/DT_*.json 悬空引用）")
    p_pcheck.add_argument("-t", "--table", help="仅检查指定逻辑表（如 Effect、Card）")
    p_pcheck.add_argument("--csv", help="导出 CSV 报告路径")
    p_pcheck.add_argument("--limit", type=int, default=20, help="每个来源表最多显示条数 (默认20)")
    p_pcheck.set_defaults(func=cmd_pcheck)

    # --- add-merge ---
    p_addmerge = sub.add_parser(
        "add-merge", help="增量合并：暂存表 → 目标表（只增只改，绝不删行）"
    )
    p_addmerge.add_argument("target", help="目标表名（部分匹配）或直接路径")
    p_addmerge.add_argument("--staging", required=True, help="暂存表 xlsx 路径")
    p_addmerge.add_argument("--apply", action="store_true",
                            help="实际写入（默认 dry-run 只打印报告）")
    p_addmerge.set_defaults(func=cmd_add_merge)

    # --- svn ---
    p_svn = sub.add_parser("svn", help="SVN 版本控制操作")
    svn_sub = p_svn.add_subparsers(dest="svn_command")

    p_svn_status = svn_sub.add_parser("status", help="查看 SVN 状态")
    p_svn_status.set_defaults(func=cmd_svn_status)

    p_svn_update = svn_sub.add_parser("update", help="SVN 更新 DataTable/Raw")
    p_svn_update.set_defaults(func=cmd_svn_update)

    p_svn_log = svn_sub.add_parser("log", help="查看 SVN 日志")
    p_svn_log.add_argument("-n", type=int, help="显示条数 (默认10)")
    p_svn_log.set_defaults(func=cmd_svn_log)

    # --- merge ---
    p_merge = sub.add_parser("merge", help="冲突检测与合并")
    merge_sub = p_merge.add_subparsers(dest="merge_command")

    p_merge_status = merge_sub.add_parser("status", help="列出冲突文件")
    p_merge_status.set_defaults(func=cmd_merge_status)

    p_merge_preview = merge_sub.add_parser("preview", help="生成合并预览文件")
    p_merge_preview.add_argument("file", help="冲突文件名（部分匹配）")
    p_merge_preview.add_argument(
        "--conflict-default",
        choices=["theirs", "mine", "empty"],
        default="theirs",
        help="冲突单元格 [最终] 行默认采用哪一侧 (默认: theirs)",
    )
    p_merge_preview.add_argument("--no-open", action="store_true", help="不自动打开 Excel")
    p_merge_preview.set_defaults(func=cmd_merge_preview)

    p_merge_resolve = merge_sub.add_parser("resolve", help="应用合并结果")
    p_merge_resolve.add_argument("file", help="冲突文件名（部分匹配）")
    p_merge_resolve.set_defaults(func=cmd_merge_resolve)

    p_merge_recover = merge_sub.add_parser("recover", help="recover changes from *_merge.xlsx")
    p_merge_recover.add_argument("file", help="original file, *_merge.xlsx, or partial name")
    p_merge_recover.add_argument("--backup", help="optional backup path; defaults to latest valid merge backup")
    p_merge_recover.set_defaults(func=cmd_merge_recover)

    p_merge_abort = merge_sub.add_parser("abort", help="放弃合并并删除临时文件")
    p_merge_abort.add_argument("file", help="冲突文件名（部分匹配）")
    p_merge_abort.set_defaults(func=cmd_merge_abort)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    if args.command == "svn":
        if not args.svn_command:
            p_svn.print_help()
            return
        args.func(args)
    elif args.command == "merge":
        if not args.merge_command:
            p_merge.print_help()
            return
        args.func(args)
    else:
        args.func(args)


if __name__ == "__main__":
    main()
