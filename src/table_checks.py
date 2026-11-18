"""
导表检查规则集。

在 manager_core.validate_table 现有的两类检查（Excel 错误值、重复 RowName）之上扩展：
  error 级：英文字段名重复、RowName 为空、类型不匹配、ID/公式引用不存在、配置文件缺失
  warning 级：中文字段名重复、表头缺失、空数据行

错误 dict 沿用 manager_core 的结构（file/sheet/row/col/col_ref/col_header/col_type/error），
额外带 level（error|warning）与 message。
"""

import manager_core as mc

ERROR = "error"
WARNING = "warning"

# H3 类型词汇表中需要校验可解析性的类型
_TYPE_INT = {"int", "int32", "int64"}
_TYPE_FLOAT = {"float", "double"}
_TYPE_BOOL = {"bool"}
_TYPE_ARRAY_INT = {"array_int", "int_array"}


def _norm(value):
    return str(value).strip() if value is not None else ""


def _norm_header(value):
    """表头归一化：空和字面量 None 都视为未定义。"""
    text = _norm(value)
    return "" if text.lower() == "none" else text


def _is_int(text):
    try:
        return float(text).is_integer()
    except (ValueError, OverflowError):
        return False


def _is_float(text):
    try:
        float(text)
        return True
    except (ValueError, OverflowError):
        return False


def _err(file, sheet, row, col, col_ref, col_header, col_type, error, level, message):
    return {
        "file": file,
        "sheet": sheet,
        "row": row,
        "col": col,
        "col_ref": col_ref,
        "col_header": col_header,
        "col_type": col_type,
        "error": error,
        "level": level,
        "message": message,
    }


# ---------------------------------------------------------------- 规则

def check_headers(file_relpath, sheet_name, ws, errors):
    """表头完整性：H1/H2 重复、有数据但 H2 缺失。"""
    import openpyxl.utils

    h1_seen, h2_seen = {}, {}
    col_has_data = {}
    for col_idx in range(1, ws.max_column + 1):
        h1 = _norm_header(ws.cell(row=1, column=col_idx).value)
        h2 = _norm_header(ws.cell(row=2, column=col_idx).value)
        if h1:
            h1_seen.setdefault(h1, []).append(col_idx)
        if h2:
            h2_seen.setdefault(h2, []).append(col_idx)
        for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
            if _norm(ws.cell(row=row_idx, column=col_idx).value):
                col_has_data[col_idx] = True
                break
        if h2 or col_idx not in col_has_data:
            continue
        errors.append(_err(
            file_relpath, sheet_name, 2, col_idx,
            f"{openpyxl.utils.get_column_letter(col_idx)}2", h1 or "—",
            _norm(ws.cell(row=3, column=col_idx).value),
            "表头缺失", WARNING,
            f"列 {openpyxl.utils.get_column_letter(col_idx)} 有数据但英文字段名（第2行）为空",
        ))
    for h1, cols in h1_seen.items():
        if len(cols) > 1:
            errors.append(_err(
                file_relpath, sheet_name, 1, cols[0],
                f"{openpyxl.utils.get_column_letter(cols[0])}1", h1, "",
                "中文字段名重复", WARNING,
                f"中文字段名 '{h1}' 重复出现，列: {cols}",
            ))
    for h2, cols in h2_seen.items():
        if len(cols) > 1:
            errors.append(_err(
                file_relpath, sheet_name, 2, cols[0],
                f"{openpyxl.utils.get_column_letter(cols[0])}2", h2, "",
                "英文字段名重复", ERROR,
                f"英文字段名 '{h2}' 重复出现，列: {cols}",
            ))


def check_rows(file_relpath, sheet_name, ws, errors):
    """RowName 必填（有数据但 A 列为空）与空数据行（仅夹在数据中间的）。"""
    row_has_data = {}
    for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
        for col_idx in range(2, ws.max_column + 1):
            if _norm(ws.cell(row=row_idx, column=col_idx).value):
                row_has_data[row_idx] = True
                break
    last_data_row = max(row_has_data, default=0)

    for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
        rowname = _norm(ws.cell(row=row_idx, column=1).value)
        has_data = row_has_data.get(row_idx, False)
        if not rowname and has_data:
            errors.append(_err(
                file_relpath, sheet_name, row_idx, 1, f"A{row_idx}", "RowName", "",
                "RowName为空", ERROR,
                f"第 {row_idx} 行有数据但 RowName（A列）为空",
            ))
        elif not rowname and not has_data and row_idx < last_data_row:
            errors.append(_err(
                file_relpath, sheet_name, row_idx, 1, f"A{row_idx}", "RowName", "",
                "空数据行", WARNING,
                f"第 {row_idx} 行为空数据行",
            ))


def check_types(file_relpath, sheet_name, ws, errors):
    """按第 3 行类型声明校验单元格值的可解析性（空值与未声明类型跳过）。"""
    import openpyxl.utils

    for col_idx in range(1, ws.max_column + 1):
        type_text = _norm(ws.cell(row=3, column=col_idx).value).lower()
        if type_text not in (_TYPE_INT | _TYPE_FLOAT | _TYPE_BOOL | _TYPE_ARRAY_INT):
            continue
        col_ref_letter = openpyxl.utils.get_column_letter(col_idx)
        col_header = _norm_header(ws.cell(row=2, column=col_idx).value) or _norm_header(
            ws.cell(row=1, column=col_idx).value
        )
        for row_idx in range(mc.DATA_START_ROW, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = _norm(value)
            if not text:
                continue
            ok = True
            if type_text in _TYPE_INT:
                ok = _is_int(text)
            elif type_text in _TYPE_FLOAT:
                ok = _is_float(text)
            elif type_text in _TYPE_BOOL:
                ok = isinstance(value, bool) or text.lower() in ("true", "false", "0", "1")
            elif type_text in _TYPE_ARRAY_INT:
                ok = all(_is_int(t) for t in text.split(",") if t.strip())
            if not ok:
                errors.append(_err(
                    file_relpath, sheet_name, row_idx, col_idx,
                    f"{col_ref_letter}{row_idx}", col_header, type_text,
                    "类型不匹配", ERROR,
                    f"值 '{text[:40]}' 无法解析为 {type_text}",
                ))


def check_dangling_refs(file_relpath, errors):
    """引用完整性：按 references.json 规则检查本文件的悬空引用。"""
    import ref_index

    if not ref_index.is_available():
        return
    index = ref_index.build_reference_index()
    for rec in index.dangling:
        if rec["src_file"] != file_relpath:
            continue
        if rec["kind"] == "formula":
            error_name = "引用目标不存在"
        else:
            error_name = "ID引用不存在"
        errors.append(_err(
            rec["src_file"], rec["src_sheet"], rec["src_row"], rec["src_col"],
            rec["src_col_ref"], rec["src_header"], "",
            error_name, rec.get("level", ERROR),
            f"[{rec['rule']}] 引用值 '{rec['value']}' 在目标表中不存在",
        ))


def check_missing_files(categories_config, errors):
    """categories.json 引用了不存在的文件。"""
    files = []
    for cat in categories_config.get("categories", {}).values():
        files.extend(cat.get("files", []))
    if categories_config.get("uncategorized"):
        files.extend(categories_config["uncategorized"].get("files", []))
    for rel in dict.fromkeys(files):
        if not (mc.RAW_DIR / rel).exists():
            errors.append(_err(
                rel, "", 0, 0, "", "", "",
                "配置文件缺失", ERROR,
                f"categories.json 引用的文件不存在: {rel}",
            ))


def run_all(file_relpath, wb):
    """对单个已打开的 workbook 运行全部表级规则（除引用完整性外的部分）。"""
    errors = []
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("#"):
            continue
        ws = wb[sheet_name]
        check_headers(file_relpath, sheet_name, ws, errors)
        check_rows(file_relpath, sheet_name, ws, errors)
        check_types(file_relpath, sheet_name, ws, errors)
    return errors
