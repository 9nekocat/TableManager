# -*- coding: utf-8 -*-
"""
DataTable Manager 核心逻辑
- 分类发现表格
- 校验 #N/A 等 Excel 错误值
- 打开文件
"""

import os
import sys
import json
import csv
from datetime import datetime
from pathlib import Path

import openpyxl

# --- 路径工具 ---
def _get_data_dir():
    """获取配置/数据目录（src/ 的同级 data/；打包后为 _MEIPASS/data）。"""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / "data"
    return Path(__file__).resolve().parent.parent / "data"


def _get_templates_dir():
    """获取前端构建产物目录（gui/dist；打包后为 _MEIPASS/gui/dist）。"""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / "gui" / "dist"
    return Path(__file__).resolve().parent.parent / "gui" / "dist"


def _find_project_root():
    """从脚本/exe位置向上查找项目根目录（包含 DataTable/Raw 或 .uproject 文件）。"""
    # 优先从可执行文件位置向上找（exe 可能放在项目任意位置）
    exe_dir = Path(sys.executable).parent.absolute() if getattr(sys, 'frozen', False) else None
    script_dir = Path(__file__).parent.absolute()

    for base in [exe_dir, script_dir]:
        if base is None:
            continue
        current = base
        for _ in range(6):
            if (current / "DataTable" / "Raw").exists():
                return current
            if list(current.glob("*.uproject")):
                return current
            current = current.parent

    # 兜底：环境变量
    env_root = os.environ.get("CATSINTHESHELL_ROOT")
    if env_root:
        return Path(env_root)

    raise FileNotFoundError(
        "无法找到项目根目录。请将 TableManager.exe 放在项目目录下，"
        "或设置环境变量 CATSINTHESHELL_ROOT 指向项目根目录。"
    )


PROJECT_ROOT = _find_project_root()
RAW_DIR = PROJECT_ROOT / "DataTable" / "Raw"
TEMP_DIR = RAW_DIR / ".temp"
PROCESSED_DIR = PROJECT_ROOT / "DataTable" / "Processed"
DATA_DIR = _get_data_dir()
CONFIG_PATH = DATA_DIR / "categories.json"
TEMPLATES_DIR = _get_templates_dir()

# Excel 错误值集合
ERROR_VALUES = {"#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!"}
DUPLICATE_ID_ERROR = "重复ID"

# 数据行起始（跳过 3 行表头: 中文名 / 英文字段名 / 类型）


DATA_START_ROW = 4


# ---------------------------------------------------------------------------
# SVN conflict artifact cleanup (shared by CLI and Web API)
# ---------------------------------------------------------------------------

def cleanup_conflict_artifacts(output_filepath):
    """Delete SVN conflict artifacts (*.rNNNN, *.mine) for a given output file.

    Returns dict: {deleted: [filenames], failed: [filenames]}
    """
    import glob
    from pathlib import Path
    output_path = Path(output_filepath)
    stem = output_path.stem
    parent = output_path.parent

    deleted = []
    failed = []

    # Clean up .r* files (e.g., Effect.xlsx.r4523)
    # glob.escape: the stem may contain glob metacharacters ([, ?, *).
    for f in parent.glob(f"{glob.escape(stem)}.*.r*"):
        try:
            os.remove(f)
            deleted.append(f.name)
        except OSError:
            failed.append(f.name)

    # Clean up .mine file
    mine_file = parent / f"{output_path.name}.mine"
    if mine_file.exists():
        try:
            os.remove(mine_file)
            deleted.append(mine_file.name)
        except OSError:
            failed.append(mine_file.name)

    return {"deleted": deleted, "failed": failed}


# ---------------------------------------------------------------------------
# 类别配置加载
# ---------------------------------------------------------------------------

def load_categories(config_path=None):
    """
    加载 categories.json，校验文件存在性。
    返回: {
        "categories": { key: {display_name, description, files: [str]} },
        "uncategorized": {display_name, description, files: [str]},
        "by_file": { relpath: category_key }   # 反向索引
    }
    """
    path = Path(config_path) if config_path else CONFIG_PATH
    if not path.exists():
        raise FileNotFoundError(f"Config not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    result = {"categories": {}, "uncategorized": None, "by_file": {}}
    all_config_files = set()

    for key, cat in raw.get("categories", {}).items():
        files = cat.get("files", [])
        result["categories"][key] = {
            "display_name": cat.get("display_name", key),
            "description": cat.get("description", ""),
            "files": files,
        }
        for fp in files:
            all_config_files.add(fp)
            result["by_file"][fp] = key

    if "uncategorized" in raw:
        result["uncategorized"] = raw["uncategorized"]
        for fp in raw["uncategorized"].get("files", []):
            all_config_files.add(fp)
            result["by_file"][fp] = "uncategorized"

    # 检查文件是否存在
    missing = [fp for fp in all_config_files if not (RAW_DIR / fp).exists()]
    if missing:
        print(f"[WARN] 以下配置中引用的文件不存在: {missing}")

    return result


# ---------------------------------------------------------------------------
# UE DataTable 名称推断
# ---------------------------------------------------------------------------

# 从 Processed/ JSON 文件名推断的映射表
_UE_NAME_MAP = {
    "CardTable.xlsx": "DT_CardDataTable",
    "CardShape.xlsx": "DT_CardShapeDataTable",
    "Effect.xlsx": "DT_RavenEffectDataTable",
    "Effectnumber.xlsx": "DT_RavenEffectNumberDataTable",
    "Target.xlsx": "DT_RavenTargetDataTable",
    "Event.xlsx": "DT_RavenEventDataTable",
    "EventStyle.xlsx": "DT_EventStyleTable",
    "EventTypePool.xlsx": "DT_EventTypePoolTable",
    "DialogEffectData.xlsx": "DT_DialogEffectDataTable",
    "DialogActionData.xlsx": "DT_DialogActionDataTable",
    "DialogConditionData.xlsx": "DT_DialogConditionDataTable",
    "ProjectileTable.xlsx": "DT_ProjectileDataTable",
    "BattleDropGroupType.xlsx": "DT_BattleDropGroupTypeTable",
    "MatchPoolConstraint.xlsx": "DT_MatchPoolConstraint",
    "FlagConfigRow.xlsx": "DT_FlagConfigRowTable",
    "GameplayTagTableRow.xlsx": "DT_GameplayTagTableRowTable",
    "StoryInteractableRow.xlsx": "DT_InteractableConfigRowTable",
    "StoryMapPointData.xlsx": "DT_StoryMapPointDataTable",
    "StoryMapRegionData.xlsx": "DT_StoryMapRegionDataTable",
    "StoryMapRoomData.xlsx": "DT_StoryMapRoomDataTable",
    "Localization/Card.xlsx": "DT_CardDataTable",
    "Localization/AttributeTable.xlsx": "DT_AttributeDataTable",
    "Localization/Buff.xlsx": "DT_RavenBuffDataTable",
    "Localization/MechaTable.xlsx": "DT_MechaDataTable",
    "Localization/Drone.xlsx": "DT_DroneDataTable",
    "Localization/Tag.xlsx": "DT_GameplayTagTable",
    "Localization/DialogData2.xlsx": "DT_DialogData2Table",
    "Localization/PvpCatTableTable.xlsx": "DT_PvpCatTableTable",
    "Localization/RavenPilotSkillData.xlsx": "DT_RavenPilotSkillDataTable",
    "Localization/ShopFilterCondition.xlsx": "DT_ShopFilterConditionTable",
    "Localization/TacticalEncounterData.xlsx": "DT_TacticalEncounterDataTable",
}


def infer_ue_name(filepath):
    """根据 xlsx 相对路径推断 UE DataTable 名称。"""
    rel = str(filepath)
    if hasattr(filepath, "as_posix"):
        rel = filepath.as_posix()
    if rel.startswith("Raw/"):
        rel = rel[4:]
    return _UE_NAME_MAP.get(rel, "")


# ---------------------------------------------------------------------------
# 表格发现
# ---------------------------------------------------------------------------

def _load_workbook_safe(filepath):
    """安全加载 workbook，失败返回 None。"""
    try:
        return openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"[ERROR] 无法打开 {filepath}: {e}")
        return None


def get_table_info(file_relpath):
    """
    获取单个 xlsx 文件的信息。
    返回 dict: file, path, category, sheets, data_sheets, row_count, col_count, ue_name, last_modified, size_bytes
    """
    fullpath = RAW_DIR / file_relpath
    stat = fullpath.stat()
    info = {
        "file": file_relpath,
        "path": fullpath,
        "category": "",
        "sheets": [],
        "data_sheets": [],
        "row_count": 0,
        "col_count": 0,
        "ue_name": "",
        "last_modified": datetime.fromtimestamp(stat.st_mtime),
        "size_bytes": stat.st_size,
    }

    wb = _load_workbook_safe(fullpath)
    if not wb:
        return info

    info["sheets"] = wb.sheetnames
    info["data_sheets"] = [s for s in wb.sheetnames if s.startswith("#")]
    info["ue_name"] = infer_ue_name(file_relpath)

    max_rows = 0
    max_cols = 0
    for sn in info["data_sheets"]:
        ws = wb[sn]
        data_rows = max(0, ws.max_row - DATA_START_ROW + 1)
        if data_rows > max_rows:
            max_rows = data_rows
        if ws.max_column > max_cols:
            max_cols = ws.max_column

    info["row_count"] = max_rows
    info["col_count"] = max_cols
    wb.close()
    return info


def discover_tables(categories_config):
    """
    扫描 Raw/ 目录下所有 .xlsx，按类别分组并返回 table_info 列表。
    返回: (tables_by_category, uncataloged_files)
    """
    by_cat = {}  # category_key -> [table_info]
    for key in categories_config["categories"]:
        by_cat[key] = []

    all_configured = set()
    for key, cat in categories_config["categories"].items():
        for fp in cat["files"]:
            all_configured.add(fp)
            if (RAW_DIR / fp).exists():
                info = get_table_info(fp)
                info["category"] = key
                by_cat[key].append(info)

    # 未分类的
    if categories_config.get("uncategorized"):
        by_cat["uncategorized"] = []
        for fp in categories_config["uncategorized"].get("files", []):
            all_configured.add(fp)
            if (RAW_DIR / fp).exists():
                info = get_table_info(fp)
                info["category"] = "uncategorized"
                by_cat["uncategorized"].append(info)

    # 扫描目录，找出未在配置中记录的文件
    uncataloged = []
    for root, dirs, files in os.walk(RAW_DIR):
        # 跳过 .temp/ 等隐藏目录（合并预览、差异导出、备份等临时产物）
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        # 跳过 Processed/Local/Macro 等不是 Raw 子目录的
        for fname in files:
            if not fname.endswith(".xlsx") or fname.startswith("~$"):
                continue
            full = Path(root) / fname
            rel = full.relative_to(RAW_DIR).as_posix()
            if rel not in all_configured:
                uncataloged.append(rel)

    return by_cat, uncataloged


# ---------------------------------------------------------------------------
# 表格搜索
# ---------------------------------------------------------------------------

def find_table(query, categories_config):
    """
    按文件名或 UE DataTable 名模糊搜索。
    返回: (file_relpath, table_info) 或 (None, None)
    """
    query_lower = query.lower().replace("dt_", "").replace("_", "")

    # 先精确匹配文件名
    all_files = []
    for key, cat in categories_config.get("categories", {}).items():
        all_files.extend(cat.get("files", []))
    if categories_config.get("uncategorized"):
        all_files.extend(categories_config["uncategorized"].get("files", []))

    # 直接文件名匹配
    for fp in all_files:
        name = Path(fp).stem.lower()
        if name == query_lower:
            return fp, get_table_info(fp)

    # 模糊匹配（文件名包含查询词）
    best = []
    for fp in all_files:
        name = Path(fp).stem.lower()
        fp_lower = fp.lower()
        ue = _UE_NAME_MAP.get(fp, "").lower()
        # 匹配文件名
        if query_lower in name or query_lower in fp_lower:
            best.append(fp)
        # 匹配 UE 名称
        elif ue and query_lower in ue:
            best.append(fp)

    if best:
        # 返回第一个匹配（最短文件名优先）
        best.sort(key=lambda x: len(x))
        fp = best[0]
        return fp, get_table_info(fp)

    return None, None


# ---------------------------------------------------------------------------
# 打开文件
# ---------------------------------------------------------------------------

def open_in_excel(file_relpath):
    """在 Excel 中打开文件。"""
    fullpath = RAW_DIR / file_relpath
    if not fullpath.exists():
        raise FileNotFoundError(f"File not found: {fullpath}")
    os.startfile(str(fullpath))


# ---------------------------------------------------------------------------
# 校验
# ---------------------------------------------------------------------------

def validate_table(file_relpath):
    """
    校验单个 xlsx 文件：Excel 错误值、重复 RowName，以及 table_checks 中的
    扩展导表检查（表头/RowName/类型/引用完整性）。
    返回: list[dict] — 每个 dict: file, sheet, row, col, col_header, col_type, error
    """
    fullpath = RAW_DIR / file_relpath
    errors = []

    wb = _load_workbook_safe(fullpath)
    if not wb:
        return errors

    for sheet_name in wb.sheetnames:
        # 仅校验 # 开头的 data sheet
        if not sheet_name.startswith("#"):
            continue

        ws = wb[sheet_name]

        # 读取表头（第 2 行: 英文字段名, 第 3 行: 类型）
        headers = {}
        types = {}
        for col_idx in range(1, ws.max_column + 1):
            header_val = ws.cell(row=2, column=col_idx).value
            type_val = ws.cell(row=3, column=col_idx).value
            headers[col_idx] = str(header_val) if header_val is not None else ""
            types[col_idx] = str(type_val) if type_val is not None else ""

        # 扫描数据行（第 4 行开始）
        rowname_rows = {}
        for row_idx in range(DATA_START_ROW, ws.max_row + 1):
            rowname_val = _normalize_cell_text(ws.cell(row=row_idx, column=1).value)
            if rowname_val:
                rowname_rows.setdefault(rowname_val, []).append(row_idx)

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_val = _normalize_cell_text(cell.value)

                if cell_val in ERROR_VALUES:
                    errors.append({
                        "file": file_relpath,
                        "sheet": sheet_name,
                        "row": row_idx,
                        "col": col_idx,
                        "col_ref": f"{openpyxl.utils.get_column_letter(col_idx)}{row_idx}",
                        "col_header": headers.get(col_idx, ""),
                        "col_type": types.get(col_idx, ""),
                        "error": cell_val,
                    })

        for rowname_val, rows in rowname_rows.items():
            if len(rows) <= 1:
                continue

            rows_text = ", ".join(str(r) for r in rows)
            for row_idx in rows:
                errors.append({
                    "file": file_relpath,
                    "sheet": sheet_name,
                    "row": row_idx,
                    "col": 1,
                    "col_ref": f"A{row_idx}",
                    "col_header": headers.get(1, "RowName"),
                    "col_type": types.get(1, ""),
                    "error": DUPLICATE_ID_ERROR,
                    "value": rowname_val,
                    "duplicate_rows": rows,
                    "message": f"重复 ID: {rowname_val}，涉及行: {rows_text}",
                })

    # 扩展导表检查（表头/RowName/类型规则），失败不影响既有检查
    try:
        import table_checks
        errors.extend(table_checks.run_all(file_relpath, wb))
    except Exception:
        pass

    wb.close()

    # 引用完整性检查（跨表，依赖 references.json 与全量引用索引）
    try:
        import table_checks
        table_checks.check_dangling_refs(file_relpath, errors)
    except Exception:
        pass
    return errors


def _normalize_cell_text(value):
    """Normalize a cell value for validation comparisons."""
    return str(value).strip() if value is not None else ""


def validate_all(categories_config, file_filter=None):
    """
    批量校验所有配置中的表格。
    file_filter: 可选的单个文件名，只校验该文件。
    返回: list[dict] 错误列表
    """
    all_files = []
    for key, cat in categories_config.get("categories", {}).items():
        all_files.extend(cat.get("files", []))
    if categories_config.get("uncategorized"):
        all_files.extend(categories_config["uncategorized"].get("files", []))

    if file_filter:
        # 支持部分文件名匹配
        matching = [f for f in all_files if file_filter.lower() in f.lower()]
        if not matching:
            # 尝试匹配不带路径的纯文件名
            matching = [f for f in all_files if file_filter.lower() in Path(f).name.lower()]
        all_files = matching

    errors = []
    if not file_filter:
        # categories.json 引用了不存在的文件
        try:
            import table_checks
            table_checks.check_missing_files(categories_config, errors)
        except Exception:
            pass
    for fp in all_files:
        fullpath = RAW_DIR / fp
        if not fullpath.exists():
            continue
        file_errors = validate_table(fp)
        errors.extend(file_errors)

    return errors


# ---------------------------------------------------------------------------
# 输出格式化
# ---------------------------------------------------------------------------

def _format_size(size_bytes):
    """人类可读的文件大小。"""
    if size_bytes < 1024:
        return f"{size_bytes}B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.0f}KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f}MB"


def _col_ref(col_idx):
    """列索引 -> Excel 列字母。"""
    return openpyxl.utils.get_column_letter(col_idx)


def format_table_list(by_cat, categories_config, category_filter=None):
    """
    格式化表格列表输出。
    """
    lines = []
    lines.append("=" * 80)
    lines.append("  DataTable Manager - 配置表一览")
    lines.append(f"  Raw: {RAW_DIR}")
    lines.append("=" * 80)

    total_files = 0
    cat_keys = list(by_cat.keys())

    for key in cat_keys:
        if category_filter and key.lower() != category_filter.lower():
            continue

        tables = by_cat.get(key, [])
        if not tables:
            continue
        total_files += len(tables)

        cat = categories_config.get("categories", {}).get(key, {})
        if key == "uncategorized":
            cat = categories_config.get("uncategorized", {})

        display = cat.get("display_name", key)
        desc = cat.get("description", "")
        header = f"[{display}]" + (f" — {desc}" if desc else "")
        lines.append(f"\n  {header}")

        for info in tables:
            loc_tag = " [本地化]" if info["file"].startswith("Localization/") else ""
            size_str = _format_size(info["size_bytes"])
            rows_str = f"{info['row_count']}行" if info["row_count"] else "—"
            cols_str = f"{info['col_count']}列" if info["col_count"] else "—"

            # 格式:   Effect.xlsx         3212行 x 16列  626KB  DT_RavenEffectDataTable
            name_tag = f"{info['file']}{loc_tag}"
            name_pad = name_tag.ljust(42)
            dims = f"{rows_str} x {cols_str}".ljust(16)
            ue = info.get("ue_name", "")
            lines.append(f"    {name_pad} {dims} {size_str:>6s}   {ue}")

    # 未编目文件
    lines.append("")
    return "\n".join(lines)


def format_validation_report(errors, csv_file=None):
    """
    格式化校验报告。
    csv_file: 如果指定，同时导出 CSV。
    """
    if not errors:
        return "[OK] 导表检查通过：未发现问题。"

    lines = []
    lines.append("=" * 80)
    lines.append("  导表检查报告 — 错误值 / 重复ID / 表头 / 类型 / 引用完整性")
    lines.append("=" * 80)

    # 按文件分组
    by_file = {}
    for e in errors:
        f = e["file"]
        by_file.setdefault(f, []).append(e)

    # 统计
    by_type = {}
    for e in errors:
        et = e["error"]
        by_type[et] = by_type.get(et, 0) + 1

    # 按文件输出
    for fname in sorted(by_file.keys()):
        file_errors = by_file[fname]
        lines.append(f"\n  [FILE] {fname} ({len(file_errors)} 个错误)")

        # 再按 sheet 分组
        by_sheet = {}
        for e in file_errors:
            by_sheet.setdefault(e["sheet"], []).append(e)

        for sname in sorted(by_sheet.keys()):
            sheet_errors = by_sheet[sname]
            lines.append(f"   [SHEET] {sname} ({len(sheet_errors)} 个)")

            for e in sheet_errors:
                col_info = f" [{e['col_header']}]" if e["col_header"] else ""
                detail = f"  {e['message']}" if e.get("message") else ""
                level_tag = " (warning)" if e.get("level") == "warning" else ""
                lines.append(
                    f"      {e['error']:>8s}  {e['col_ref']:>6s}  "
                    f"第{e['row']}行 第{e['col']:>3d}列{col_info}{detail}{level_tag}"
                )

    # 汇总
    lines.append(f"\n{'─' * 60}")
    lines.append(f"  总计: {len(errors)} 个错误, {len(by_file)} 个文件")
    for et, cnt in sorted(by_type.items(), key=lambda x: -x[1]):
        lines.append(f"    {et}: {cnt}")

    report = "\n".join(lines)

    # CSV 导出
    if csv_file:
        _export_errors_csv(errors, csv_file)
        report += f"\n\n[CSV] 报告已导出至: {csv_file}"

    return report


def _export_errors_csv(errors, csv_path):
    """导出错误列表为 CSV。"""
    if not errors:
        return
    fieldnames = [
        "file", "sheet", "row", "col", "col_ref", "col_header", "col_type",
        "error", "level", "value", "duplicate_rows", "message",
    ]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(errors)


# ---------------------------------------------------------------------------
# RowName 提取工具 (供 merge_engine 使用)
# ---------------------------------------------------------------------------

def get_rowname_map(file_relpath):
    """
    读取 xlsx 文件，提取每个 data sheet 的 RowName 映射。
    返回: {sheet_name: {row_index: rowname_value}}
    """
    fullpath = RAW_DIR / file_relpath
    if not fullpath.exists():
        return {}

    wb = _load_workbook_safe(fullpath)
    if not wb:
        return {}

    result = {}
    for sn in wb.sheetnames:
        if not sn.startswith("#"):
            continue
        ws = wb[sn]
        sheet_map = {}
        for row_idx in range(DATA_START_ROW, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=1).value
            if val is not None:
                sheet_map[row_idx] = str(val).strip()
        result[sn] = sheet_map

    wb.close()
    return result


def format_table_info(info):
    """格式化单表详情输出。"""
    lines = []
    lines.append("=" * 60)
    lines.append(f"  {info['file']}")
    lines.append("=" * 60)
    lines.append(f"  路径:       {info['path']}")
    lines.append(f"  类别:       {info['category']}")
    lines.append(f"  UE 表名:    {info['ue_name'] or '(未映射)'}")
    lines.append(f"  最后修改:   {info['last_modified'].strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"  文件大小:   {_format_size(info['size_bytes'])}")
    lines.append(f"  Sheets:     {', '.join(info['sheets'])}")
    lines.append(f"  DataSheets: {', '.join(info['data_sheets']) if info['data_sheets'] else '(无 # 开头的 data sheet)'}")
    lines.append(f"  数据行数:   {info['row_count']}")
    lines.append(f"  列数:       {info['col_count']}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 类别配置修改
# ---------------------------------------------------------------------------

def move_file_to_category(filepath, to_category, config_path=None):
    """Move a file from its current category to a new one in categories.json.

    Uses atomic write (write temp + os.replace) to prevent config corruption.
    Returns: {"ok": True, "from": old_category, "to": to_category}
    """
    import json
    import tempfile

    path = Path(config_path) if config_path else CONFIG_PATH
    with open(path, "r", encoding="utf-8") as f:
        config = json.load(f)

    old_category = None

    # Remove from all categories + uncategorized
    for cat_key, cat in config.get("categories", {}).items():
        files = cat.get("files", [])
        if filepath in files:
            files.remove(filepath)
            old_category = cat_key
            break

    if old_category is None and "uncategorized" in config:
        files = config["uncategorized"].get("files", [])
        if filepath in files:
            files.remove(filepath)
            old_category = "uncategorized"

    if old_category is None:
        raise ValueError(f"File not found in any category: {filepath}")

    # Add to target category
    if to_category == "uncategorized":
        config.setdefault("uncategorized", {"display_name": "未分类", "files": []})
        config["uncategorized"]["files"].append(filepath)
    else:
        if to_category not in config.get("categories", {}):
            raise ValueError(f"Target category not found: {to_category}")
        config["categories"][to_category]["files"].append(filepath)

    # Atomic write: temp file + os.replace
    tmp = tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", delete=False, dir=path.parent)
    try:
        json.dump(config, tmp, ensure_ascii=False, indent=2)
        tmp.close()
        os.replace(tmp.name, path)
    except Exception:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
        raise

    return {"ok": True, "from": old_category, "to": to_category}
