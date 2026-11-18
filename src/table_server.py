# -*- coding: utf-8 -*-
"""
DataTable Manager Web Server
Flask REST API + Web GUI for table management.

Usage:
    python table_server.py              # Native window mode (pywebview)
    python table_server.py --web        # Web browser mode
    python table_server.py --port 8080  # Custom port
    python table_server.py --no-open    # Don't auto-open browser
"""

import argparse
import os
import sys
import socket
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

# 源码运行时把 src/ 自身加入模块搜索路径（打包后由 PyInstaller 处理）
sys.path.insert(0, str(Path(__file__).resolve().parent))

from manager_core import (
    RAW_DIR,
    TEMP_DIR,
    DATA_DIR,
    TEMPLATES_DIR,
    DATA_START_ROW,
    load_categories,
    discover_tables,
    find_table,
    open_in_excel,
    get_table_info,
    validate_all,
    validate_table,
    move_file_to_category,
)
from svn_client import SvnClient, SvnError, ensure_svn_available
from merge_engine import generate_merge_preview
from merge_workflow import (
    MergeWorkflowError,
    apply_merge_preview_transactional,
    recover_merge_preview_transactional,
    find_merge_preview,
    original_name_from_preview,
    infer_theirs_path,
)
from diff_engine import generate_diff, export_diff_to_excel
import toolbox  # 妙妙工具箱（Effect 重算修复 / 查 #N/A）

app = Flask(__name__, static_folder=None)
_categories_config = None
_tables_cache = None

# Serialize destructive filesystem/SVN operations (merge resolve/recover/abort,
# svn update) — Flask runs threaded by default.
_merge_lock = threading.Lock()


def _safe_join(base, name):
    """Join `name` under `base`, rejecting path traversal.
    Returns a resolved Path inside `base`, or None if it escapes."""
    try:
        p = (Path(base) / name).resolve()
        p.relative_to(Path(base).resolve())
        return p
    except (ValueError, OSError):
        return None


def invalidate_config_cache(include_categories=False):
    """Invalidate cached table/category data after filesystem changes."""
    global _categories_config, _tables_cache
    _tables_cache = None
    if include_categories:
        _categories_config = None


def get_config(refresh_tables=False):
    global _categories_config, _tables_cache
    if _categories_config is None:
        config = load_categories()
        _categories_config = config
    if refresh_tables or _tables_cache is None:
        tables, _uncat = discover_tables(_categories_config)
        _tables_cache = tables
    return _categories_config, _tables_cache


# ====== 前端页面 ======
@app.route("/")
def index():
    return send_from_directory(str(TEMPLATES_DIR), "index.html")

@app.route("/assets/<path:filename>")
def serve_assets(filename):
    assets_dir = TEMPLATES_DIR / "assets"
    return send_from_directory(str(assets_dir), filename)

@app.route("/art/<path:filename>")
def serve_art(filename):
    art_dir = TEMPLATES_DIR / "art"
    return send_from_directory(str(art_dir), filename)


# ====== API: 表格列表 ======
@app.route("/api/tables")
def api_tables():
    config, by_cat = get_config(refresh_tables=True)
    result = []

    for cat_key, tables in by_cat.items():
        cat_info = config.get("categories", {}).get(cat_key, {})
        if cat_key == "uncategorized":
            cat_info = config.get("uncategorized", {})

        cat_data = {
            "key": cat_key,
            "display_name": cat_info.get("display_name", cat_key),
            "description": cat_info.get("description", ""),
            "tables": [],
        }

        for t in tables:
            cat_data["tables"].append({
                "file": t["file"],
                "filename": Path(t["file"]).name,
                "row_count": t["row_count"],
                "col_count": t["col_count"],
                "size_bytes": t["size_bytes"],
                "size_fmt": _fmt_size(t["size_bytes"]),
                "ue_name": t.get("ue_name", ""),
                "last_modified": t["last_modified"].strftime("%Y-%m-%d %H:%M"),
                "sheets": t.get("sheets", []),
                "data_sheets": t.get("data_sheets", []),
                "is_localization": t["file"].startswith("Localization/"),
            })

        result.append(cat_data)

    # 排序：有表的在前，按 display_name 排
    result.sort(key=lambda c: (len(c["tables"]) == 0, c["display_name"]))

    return jsonify(result)


# ====== API: 环境检查 ======
@app.route("/api/env/status")
def api_env_status():
    svn = ensure_svn_available(auto_install=True, raise_on_error=False)
    return jsonify({
        "ok": svn.get("installed", False),
        "svn": svn,
    })


# ====== API: 表详情 ======
@app.route("/api/tables/<path:name>")
def api_table_info(name):
    config, _ = get_config()
    fp, info = find_table(name, config)
    if not fp:
        # 尝试直接路径
        direct = RAW_DIR / name
        if direct.exists():
            info = get_table_info(name)
            info["category"] = config.get("by_file", {}).get(name, "未分类")
        else:
            return jsonify({"error": f"找不到表格: {name}"}), 404

    return jsonify({
        "file": info["file"],
        "filename": Path(info["file"]).name,
        "category": info.get("category", ""),
        "row_count": info["row_count"],
        "col_count": info["col_count"],
        "size_bytes": info["size_bytes"],
        "size_fmt": _fmt_size(info["size_bytes"]),
        "ue_name": info.get("ue_name", ""),
        "last_modified": info["last_modified"].strftime("%Y-%m-%d %H:%M"),
        "sheets": info.get("sheets", []),
        "data_sheets": info.get("data_sheets", []),
    })


# ====== API: 校验 ======
@app.route("/api/check", methods=["POST"])
def api_check():
    config, _ = get_config()
    data = request.get_json(silent=True) or {}
    file_filter = data.get("file")

    errors = validate_all(config, file_filter=file_filter)

    # 按文件汇总
    by_file = {}
    by_type = {}
    for e in errors:
        f = e["file"]
        by_file.setdefault(f, {"file": f, "count": 0, "by_type": {}, "by_sheet": {}, "errors": []})
        by_file[f]["count"] += 1
        by_file[f]["by_type"][e["error"]] = by_file[f]["by_type"].get(e["error"], 0) + 1
        by_type[e["error"]] = by_type.get(e["error"], 0) + 1
        entry = {
            "row": e["row"],
            "col": e["col"],
            "col_ref": e["col_ref"],
            "col_header": e.get("col_header", ""),
            "col_type": e.get("col_type", ""),
            "error": e["error"],
            "level": e.get("level", "error"),
            "value": e.get("value", ""),
            "duplicate_rows": e.get("duplicate_rows", []),
            "message": e.get("message", ""),
        }
        by_file[f]["by_sheet"].setdefault(e["sheet"], []).append(entry)
        by_file[f]["errors"].append(entry)

    return jsonify({
        "total_errors": len(errors),
        "files_affected": len(by_file),
        "by_type": by_type,
        "files": [by_file[f] for f in sorted(by_file.keys())],
    })


# ====== API: 引用关系查找 ======
def _serialize_ref_record(r):
    return {
        "src_file": r["src_file"],
        "src_sheet": r["src_sheet"],
        "src_row": r["src_row"],
        "src_col_ref": r["src_col_ref"],
        "src_header": r["src_header"],
        "kind": r["kind"],
        "rule": r["rule"],
        "value": r["value"],
        "level": r.get("level", "error"),
        "targets": [
            {"file": t[0], "sheet": t[1], "row": t[2]} for t in r["targets_hit"]
        ],
    }


@app.route("/api/refs")
def api_refs():
    import ref_index

    if not ref_index.is_available():
        return jsonify({"error": "references.json 未配置，引用查找不可用"}), 400

    config, _ = get_config()
    target = request.args.get("file", "")
    fp, _ = find_table(target, config)
    if not fp:
        direct = RAW_DIR / target
        if direct.exists():
            fp = target
        else:
            return jsonify({"error": f"找不到表格: {target}"}), 404

    sheet = request.args.get("sheet") or None
    row = request.args.get("row", type=int)

    index = ref_index.build_reference_index()
    records = index.find_references(fp, sheet=sheet, row=row)
    return jsonify({
        "ok": True,
        "target": fp,
        "total": len(records),
        "references": [_serialize_ref_record(r) for r in records],
    })


@app.route("/api/refs/dangling")
def api_refs_dangling():
    import ref_index

    if not ref_index.is_available():
        return jsonify({"error": "references.json 未配置，引用查找不可用"}), 400

    src_filter = request.args.get("file") or None
    index = ref_index.build_reference_index()
    records = [r for r in index.dangling if not src_filter or r["src_file"] == src_filter]
    return jsonify({
        "ok": True,
        "total": len(records),
        "references": [_serialize_ref_record(r) for r in records],
    })


# ====== API: 导出数据检查（Processed） ======
@app.route("/api/pcheck")
def api_pcheck():
    import processed_checks

    if not processed_checks.is_available():
        return jsonify({"error": "processed_checks.json 未配置或 Processed 目录不存在"}), 400

    table_filter = request.args.get("table") or None
    result = processed_checks.run_checks(table_filter=table_filter)

    by_file = {}
    by_type = {}
    for r in result["dangling"]:
        f = r["src_file"]
        by_file.setdefault(f, {"file": f, "count": 0, "by_type": {}, "by_sheet": {}, "errors": []})
        by_file[f]["count"] += 1
        by_file[f]["by_type"][r["rule"]] = by_file[f]["by_type"].get(r["rule"], 0) + 1
        by_type[r["rule"]] = by_type.get(r["rule"], 0) + 1
        entry = {
            "row": r["src_row"],
            "col": 0,
            "col_ref": r["src_col_ref"],
            "col_header": r["src_header"],
            "col_type": "",
            "error": r["rule"],
            "level": r["level"],
            "value": r["value"],
            "duplicate_rows": [],
            "message": f"[{r['rule']}] 引用值 '{r['value']}' 在目标表中不存在",
        }
        by_file[f]["by_sheet"].setdefault("", []).append(entry)
        by_file[f]["errors"].append(entry)

    return jsonify({
        "total_errors": result["stats"]["total_dangling"],
        "total_records": result["stats"]["total_records"],
        "files_affected": len(by_file),
        "by_type": by_type,
        "skipped": result["skipped"],
        "files": [by_file[f] for f in sorted(by_file.keys())],
    })


# ====== API: 打开文件 ======
@app.route("/api/open/<path:name>", methods=["POST"])
def api_open(name):
    config, _ = get_config()
    fp, info = find_table(name, config)
    if not fp:
        # Check TEMP_DIR first (merge previews, diff exports)
        temp_path = _safe_join(TEMP_DIR, name)
        if temp_path is None:
            return jsonify({"error": f"非法路径: {name}"}), 400
        if temp_path.exists():
            os.startfile(str(temp_path))
            return jsonify({"ok": True, "file": str(temp_path)})
        # Fall back to RAW_DIR
        direct = _safe_join(RAW_DIR, name)
        if direct is None:
            return jsonify({"error": f"非法路径: {name}"}), 400
        if direct.exists():
            fp = name
        else:
            return jsonify({"error": f"找不到表格: {name}"}), 404

    try:
        open_in_excel(fp)
        return jsonify({"ok": True, "file": fp})
    except MergeWorkflowError as e:
        invalidate_config_cache()
        return jsonify(e.to_dict()), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: SVN 状态 ======
@app.route("/api/svn/status")
def api_svn_status():
    svn = SvnClient()
    try:
        is_wc = svn.is_working_copy()
    except SvnError as e:
        return jsonify({"error": str(e)}), 500

    if not is_wc:
        return jsonify({"error": "DataTable/Raw 不是 SVN 工作副本"}), 400

    try:
        items = svn.status()
        conflicts = [i for i in items if i["conflicted"]]
        modified = [i for i in items if i["status"] == "M"]
        added = [i for i in items if i["status"] == "A"]

        return jsonify({
            "clean": len(items) == 0,
            "total_changes": len(items),
            "conflicted": len(conflicts),
            "conflicted_files": [{"path": i["path"], "status": i["status"]} for i in conflicts],
            "modified_files": [{"path": i["path"], "status": i["status"]} for i in modified],
            "added_files": [{"path": i["path"], "status": i["status"]} for i in added],
            "all": [{"path": i["path"], "status": i["status"], "conflicted": i["conflicted"]} for i in items],
        })
    except SvnError as e:
        return jsonify({"error": str(e)}), 500


# ====== API: SVN 更新 ======
@app.route("/api/svn/update", methods=["POST"])
def api_svn_update():
    svn = SvnClient()
    try:
        is_wc = svn.is_working_copy()
    except SvnError as e:
        return jsonify({"error": str(e)}), 500

    if not is_wc:
        return jsonify({"error": "DataTable/Raw 不是 SVN 工作副本"}), 400

    try:
        with _merge_lock:
            summary = svn.update()
        invalidate_config_cache()
        return jsonify({
            "revision": summary["revision"],
            "updated": summary["updated"],
            "merged": summary["merged"],
            "conflicted": summary["conflicted"],
            "deleted": summary["deleted"],
            "added": summary["added"],
            "raw": [l for l in summary["raw"] if l.strip()],
        })
    except SvnError as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 合并状态 ======
@app.route("/api/merge/status")
def api_merge_status():
    # Merge features work by scanning files, no svn.exe needed
    svn = SvnClient()
    conflicts = svn.find_conflict_files()
    result = []
    seen_previews = set()
    for base_rel, entry in conflicts.items():
        preview_path = TEMP_DIR / base_rel.replace(".xlsx", "_merge.xlsx")
        seen_previews.add(preview_path.resolve())
        result.append({
            "file": base_rel,
            "mine_path": entry.get("mine", ""),
            "theirs_path": entry.get("theirs", ""),
            "base_path": entry.get("base", ""),
            "preview_exists": preview_path.exists(),
            "preview_name": str(preview_path.relative_to(TEMP_DIR)).replace("\\", "/"),
            "preview_path": str(preview_path),
        })

    recoverable = []
    if TEMP_DIR.exists():
        for preview_path in TEMP_DIR.rglob("*_merge.xlsx"):
            if preview_path.resolve() in seen_previews:
                continue
            original_name = original_name_from_preview(preview_path)
            recoverable.append({
                "file": original_name,
                "preview_exists": True,
                "preview_name": str(preview_path.relative_to(TEMP_DIR)).replace("\\", "/"),
                "preview_path": str(preview_path),
            })

    return jsonify({
        "total": len(result),
        "conflicts": result,
        "recoverable_previews": recoverable,
    })


# ====== API: 生成合并预览 ======
@app.route("/api/merge/preview/<path:name>", methods=["POST"])
def api_merge_preview(name):
    data = request.get_json(silent=True) or {}
    conflict_default = data.get("conflict_default") or data.get("conflictDefault") or "theirs"
    svn = SvnClient()
    conflicts = svn.find_conflict_files()
    target = None
    for base_rel, entry in conflicts.items():
        if name.lower() in base_rel.lower():
            target = (base_rel, entry)
            break

    if not target:
        return jsonify({"error": f"未找到冲突文件: {name}"}), 404

    base_rel, entry = target
    mine_path = entry.get("mine")
    theirs_path = entry.get("theirs")
    base_path = entry.get("base") or mine_path

    if not mine_path or not theirs_path:
        return jsonify({"error": "缺少必要的文件版本"}), 400

    output_path = TEMP_DIR / base_rel.replace(".xlsx", "_merge.xlsx")
    os.makedirs(output_path.parent, exist_ok=True)

    try:
        result = generate_merge_preview(
            base_path, mine_path, theirs_path, str(output_path),
            conflict_default=conflict_default,
        )
        os.startfile(str(output_path))
        return jsonify({
            "ok": True,
            "preview_path": str(result),
            "file": base_rel,
            "conflict_default": conflict_default,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 应用合并 ======
@app.route("/api/merge/resolve/<path:name>", methods=["POST"])
def api_merge_resolve(name):
    with _merge_lock:
        return _do_merge_resolve(name)


def _do_merge_resolve(name):
    preview_path = find_merge_preview(name)

    if not preview_path or not preview_path.exists():
        return jsonify({"error": f"找不到合并预览文件: {preview_path}"}), 404

    # Original file is in RAW_DIR (strip _merge suffix from preview name)
    original_name = original_name_from_preview(preview_path)
    output_path = RAW_DIR / original_name

    # Get theirs_path from request body or infer from conflict files
    data = request.get_json(silent=True) or {}
    theirs_path = data.get("theirs_path")
    if not theirs_path:
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

        invalidate_config_cache()

        return jsonify(result)
    except MergeWorkflowError as e:
        invalidate_config_cache()
        return jsonify(e.to_dict()), 500
    except Exception as e:
        # Keep lock on failure — user must resolve manually
        invalidate_config_cache()
        return jsonify({
            "error": str(e),
        }), 500


# ====== API: 放弃合并 ======
@app.route("/api/merge/recover/<path:name>", methods=["POST"])
def api_merge_recover(name):
    with _merge_lock:
        return _do_merge_recover(name)


def _do_merge_recover(name):
    preview_path = find_merge_preview(name)
    if not preview_path or not preview_path.exists():
        return jsonify({"error": f"找不到合并预览文件: {name}"}), 404

    original_name = original_name_from_preview(preview_path)
    output_path = RAW_DIR / original_name
    data = request.get_json(silent=True) or {}
    theirs_path = data.get("theirs_path") or infer_theirs_path(original_name)
    backup_path = data.get("backup_path")

    try:
        result = recover_merge_preview_transactional(
            preview_path,
            output_path,
            backup_path=backup_path,
            theirs_path=theirs_path,
        )
        invalidate_config_cache()
        return jsonify(result)
    except MergeWorkflowError as e:
        invalidate_config_cache()
        return jsonify(e.to_dict()), 500
    except Exception as e:
        invalidate_config_cache()
        return jsonify({"error": str(e)}), 500


@app.route("/api/merge/abort/<path:name>", methods=["POST"])
def api_merge_abort(name):
    with _merge_lock:
        return _do_merge_abort(name)


def _do_merge_abort(name):
    preview_path = None
    if name.endswith("_merge.xlsx"):
        preview_path = _safe_join(TEMP_DIR, name)
    else:
        preview_path = _safe_join(TEMP_DIR, name.replace(".xlsx", "_merge.xlsx"))

    if preview_path is None:
        return jsonify({"error": f"非法路径: {name}"}), 400

    if not preview_path.exists():
        return jsonify({"error": f"找不到合并预览文件: {preview_path}"}), 404

    try:
        Path(preview_path).unlink()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 文件历史版本列表 ======
@app.route("/api/diff/log/<path:name>")
def api_diff_log(name):
    svn = SvnClient()
    try:
        entries = svn.log_file(name, limit=30)
        return jsonify({"ok": True, "entries": entries})
    except SvnError as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 版本差异比对 ======
@app.route("/api/diff/compare", methods=["POST"])
def api_diff_compare():
    data = request.get_json(silent=True) or {}
    file_path = data.get("file")
    revision = data.get("revision")

    if not file_path or not revision:
        return jsonify({"error": "缺少 file 或 revision 参数"}), 400

    if _safe_join(RAW_DIR, file_path) is None:
        return jsonify({"error": f"非法路径: {file_path}"}), 400

    try:
        result = generate_diff(file_path, revision)
        return jsonify({"ok": True, **result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 导出差异到 Excel ======
@app.route("/api/diff/export", methods=["POST"])
def api_diff_export():
    data = request.get_json(silent=True) or {}
    file_path = data.get("file")
    revision = data.get("revision")

    if not file_path or not revision:
        return jsonify({"error": "缺少 file 或 revision 参数"}), 400

    if _safe_join(RAW_DIR, file_path) is None:
        return jsonify({"error": f"非法路径: {file_path}"}), 400

    try:
        output = export_diff_to_excel(file_path, revision)
        try:
            os.startfile(output)
        except Exception:
            pass
        return jsonify({"ok": True, "output": output})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 移动文件到其他分类 ======
@app.route("/api/categories/move", methods=["POST"])
def api_categories_move():
    data = request.get_json(silent=True) or {}
    file_path = data.get("file")
    to_category = data.get("to_category")

    if not file_path or not to_category:
        return jsonify({"error": "缺少 file 或 to_category 参数"}), 400

    try:
        result = move_file_to_category(file_path, to_category)
        invalidate_config_cache(include_categories=True)
        return jsonify(result)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: SVN 提交 ======
@app.route("/api/svn/commit", methods=["POST"])
def api_svn_commit():
    data = request.get_json(silent=True) or {}
    message = (data.get("message") or "").strip()
    files = data.get("files") or []

    if not message:
        return jsonify({"error": "commit message 不能为空"}), 400
    if not files:
        return jsonify({"error": "请选择要提交的文件"}), 400

    svn = SvnClient()
    try:
        result = svn.commit(files, message)
        invalidate_config_cache()
        return jsonify(result)
    except SvnError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: SVN 远端更新检测 ======
@app.route("/api/svn/check-updates")
def api_svn_check_updates():
    svn = SvnClient()
    try:
        result = svn.check_updates()
        return jsonify(result)
    except SvnError as e:
        return jsonify({"error": str(e), "has_update": False}), 500
    except Exception as e:
        return jsonify({"error": str(e), "has_update": False}), 500


# ====== API: 智能差异比对导出 Excel ======
@app.route("/api/diff/smart-export/<path:name>")
def api_diff_smart_export(name):
    from diff_engine import export_diff_to_excel
    from svn_client import SvnClient, SvnError

    if _safe_join(RAW_DIR, name) is None:
        return jsonify({"error": "非法路径"}), 400

    svn = SvnClient()
    try:
        try:
            items = svn.status(paths=[name])
        except SvnError:
            return jsonify({"error": "无法获取 SVN 状态"}), 500

        is_modified = any(i["status"] == "M" for i in items)
        revision = "HEAD" if is_modified else svn.get_file_prev_rev(name)

        if not revision:
            return jsonify({"error": "该文件无历史版本可供比对"}), 404

        output = export_diff_to_excel(name, revision)
        try:
            os.startfile(output)
        except Exception:
            pass
        return jsonify({"ok": True, "output": output, "revision": str(revision)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ====== API: 妙妙工具箱 ======
def _toolbox_resolve(name):
    """把前端传入的表名/相对路径解析为 Raw 下的绝对路径（防路径穿越）"""
    p = _safe_join(RAW_DIR, name)
    if p is not None:
        return str(p)
    # 允许传绝对路径（仅限 Raw 目录内）
    try:
        abs_p = Path(name).resolve()
        abs_p.relative_to(Path(RAW_DIR).resolve())
        return str(abs_p)
    except (ValueError, OSError):
        return None


@app.route("/api/toolbox/status")
def api_toolbox_status():
    """妙妙工具箱状态：Effect 当前错误缓存统计 + 占用状态"""
    try:
        return jsonify(toolbox.effect_status())
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/toolbox/effect/recalc", methods=["POST"])
def api_toolbox_effect_recalc():
    """执行 Effect #N/A 修复重算（备份 → COM 读回写+OR 保护 → CalculateFull → 保存 → 校验）"""
    data = request.get_json(silent=True) or {}
    try:
        result = toolbox.effect_recalc_run(
            path=data.get("path") or None,
            backup=data.get("backup", True),
        )
        code = 200 if result.get("ok") else 400
        return jsonify(result), code
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/toolbox/check-na", methods=["POST"])
def api_toolbox_check_na():
    """查 #N/A（通用）：统计任意 Raw 表的公式错误缓存并列出明细"""
    data = request.get_json(silent=True) or {}
    name = data.get("name") or data.get("path")
    if not name:
        return jsonify({"ok": False, "error": "缺少 name 参数"}), 400
    path = _toolbox_resolve(name)
    if not path:
        return jsonify({"ok": False, "error": "非法路径: %s" % name}), 400
    try:
        return jsonify(toolbox.check_na(path))
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def _fmt_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.0f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def _find_free_port(start=5000, max_attempts=50):
    """Find a free TCP port starting from `start`, incrementing on conflict."""
    for offset in range(max_attempts):
        port = start + offset
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError(f"Cannot find free port in range {start}-{start + max_attempts}")


def _start_native(port):
    """Start Flask in a background thread and open a native pywebview window."""
    url = f"http://127.0.0.1:{port}"

    # Try importing pywebview BEFORE starting the Flask thread — otherwise the
    # fallback path would run a second Flask instance on the same port.
    try:
        import webview
    except ImportError:
        print("[TableManager] pywebview not installed. Falling back to web mode.")
        print(f"[TableManager] 启动服务器: {url}")
        print(f"[TableManager] Raw: {RAW_DIR}")
        webbrowser.open(url)
        app.run(host="127.0.0.1", port=port, debug=False)
        sys.exit(0)

    def run_flask():
        app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)

    t = threading.Thread(target=run_flask, daemon=True)
    t.start()

    print(f"[TableManager] 启动原生窗口: {url}")
    print(f"[TableManager] Raw: {RAW_DIR}")
    webview.create_window(
        "DataTable Manager — catsintheshell",
        url,
        width=1280,
        height=860,
        min_size=(400, 500),
    )
    webview.start()
    # webview.start() blocks until window closes; Flask daemon thread exits with process
    sys.exit(0)


def main():
    parser = argparse.ArgumentParser(description="DataTable Manager Web Server")
    parser.add_argument("--port", type=int, default=0, help="端口 (默认: 动态分配)")
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器")
    parser.add_argument("--web", action="store_true", help="纯 Web 模式 (不启动原生窗口)")
    args = parser.parse_args()

    env = ensure_svn_available(auto_install=True, raise_on_error=False)
    if env.get("installed"):
        if env.get("install_attempted"):
            print(f"[TableManager] SVN 命令行工具已自动安装: {env.get('exe')}")
    else:
        print(f"[TableManager] SVN 环境未就绪: {env.get('error')}")

    port = args.port if args.port else _find_free_port()
    url = f"http://127.0.0.1:{port}"

    if args.web:
        if not args.no_open:
            webbrowser.open(url)
        print(f"[TableManager] 启动服务器: {url}")
        print(f"[TableManager] Raw: {RAW_DIR}")
        app.run(host="127.0.0.1", port=port, debug=False)
    else:
        _start_native(port)


if __name__ == "__main__":
    main()
