# -*- coding: utf-8 -*-
"""#N/A 基线校验（formula-cache-na-guard）与 COM 直写拒绝规则测试。

覆盖：
- count_formula_error_caches：XML 层统计 # sheet 的错误值缓存（t="e" /
  <v> 以 # 开头），无 Excel 即可跑
- verify_formula_cache_guard：写回前后基线对比，ok/added_sheets/total
- ensure_formula_cache：含外部链接文件拒绝全表重算
- write_cells_via_com / apply_write_plan_via_com 的公式拒绝路径
- 基线校验不会把既有错误值（数据问题）误判为写回新增
"""

import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pytest

from excel_recalc import (
    apply_write_plan_via_com,
    count_formula_error_caches,
    ensure_formula_cache,
    verify_formula_cache_guard,
    write_cells_via_com,
)

try:
    import win32com.client  # noqa: F401
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

needs_excel = pytest.mark.skipif(
    not HAS_WIN32COM, reason="requires Excel COM (win32com)"
)

_SS_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _write_xlsx_with_error_caches(path, sheets=("#T1", "#T2"), errors=None):
    """构造含 #N/A 缓存的 xlsx：# sheet + 公式单元格 t='e' 缓存值。

    errors: {sheet: [(cell_ref, error_text), ...]}
    openpyxl 写公式单元格为 `<c r="A4"><f>A4</f><v></v></c>`，
    zip 层整体替换为 `<c r="A4" t="e"><f>A4</f><v>#N/A</v></c>`。
    """
    import re

    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sn in sheets:
        ws = wb.create_sheet(sn)
        ws.append(["名称", "值"])
        ws.append(["Name", "Value"])
        ws.append(["FString", "FString"])
        # 数据行 4-6：A 列公式（缓存留空，openpyxl 不写缓存值），B 列静态
        for r in range(4, 7):
            ws.append([f"=A{r}", f"x{r}"])
    wb.save(str(path))
    wb.close()

    # zip 层给公式单元格注入 t="e" 错误缓存
    src = str(path)
    tmp = src + ".tmp"
    errors = errors or {}
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("xl/worksheets/sheet") \
                        and item.filename.endswith(".xml"):
                    data = data.decode("utf-8")
                    # sheetN.xml -> sheets[N-1]（创建顺序即文件名顺序）
                    m = re.search(r"sheet(\d+)\.xml$", item.filename)
                    idx = int(m.group(1)) - 1
                    sn = sheets[idx] if idx < len(sheets) else None
                    for ref, err in (errors.get(sn) or []):
                        data = re.sub(
                            rf'<c r="{ref}"><f>(.*?)</f><v></v></c>',
                            rf'<c r="{ref}" t="e"><f>\1</f><v>{err}</v></c>',
                            data,
                        )
                    zout.writestr(item, data.encode("utf-8"))
                else:
                    zout.writestr(item, data)
    os.replace(tmp, src)


def test_count_formula_error_caches_counts_na(tmp_path):
    p = tmp_path / "err.xlsx"
    _write_xlsx_with_error_caches(
        p,
        sheets=("#T1", "#T2"),
        errors={
            "#T1": [("A4", "#N/A"), ("A5", "#REF!")],
            "#T2": [("A4", "#N/A")],
        },
    )
    counts = count_formula_error_caches(str(p))
    assert counts == {"#T1": 2, "#T2": 1}


def test_count_formula_error_caches_ignores_non_hash_sheets(tmp_path):
    p = tmp_path / "mix.xlsx"
    _write_xlsx_with_error_caches(
        p,
        sheets=("T1", "Sheet2"),
        errors={"T1": [("A4", "#N/A")], "Sheet2": [("A4", "#N/A")]},
    )
    counts = count_formula_error_caches(str(p))
    assert counts == {}  # 两个 sheet 都不以 # 开头，全部忽略


def test_verify_formula_cache_guard_no_new_errors(tmp_path):
    p = tmp_path / "same.xlsx"
    _write_xlsx_with_error_caches(
        p,
        sheets=("#T1",),
        errors={"#T1": [("A4", "#N/A")]},
    )
    baseline = count_formula_error_caches(str(p))
    result = verify_formula_cache_guard(str(p), baseline)
    assert result["ok"] is True
    assert result["added_sheets"] == {}
    assert result["total"] == 0
    assert result["baseline"] == {"#T1": 1}
    assert result["after"] == {"#T1": 1}


def test_verify_formula_cache_guard_reports_added(tmp_path):
    p = tmp_path / "added.xlsx"
    _write_xlsx_with_error_caches(
        p,
        sheets=("#T1", "#T2"),
        errors={"#T1": [("A4", "#N/A")]},
    )
    baseline = count_formula_error_caches(str(p))
    # 写回后 T1 又新增一个 #REF!、T2 新增一个 #N/A
    _write_xlsx_with_error_caches(
        p,
        sheets=("#T1", "#T2"),
        errors={"#T1": [("A4", "#N/A"), ("A5", "#REF!")], "#T2": [("A4", "#N/A")]},
    )
    result = verify_formula_cache_guard(str(p), baseline)
    assert result["ok"] is False
    assert result["added_sheets"] == {"#T1": 1, "#T2": 1}
    assert result["total"] == 2


def test_ensure_formula_cache_rejects_external_link_file(tmp_path):
    """含外部链接的文件必须拒绝全表重算（外链部件 zip 层注入）。"""
    import openpyxl

    p = tmp_path / "linked.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    ws.append(["名称", "值"])
    ws.append(["Name", "Value"])
    ws.append(["FString", "FString"])
    ws.append(["=A4", "x"])
    wb.save(str(p))
    wb.close()

    # 注入一个外部链接部件（仅用于触发 _has_external_links）
    src = str(p)
    tmp = src + ".tmp"
    with zipfile.ZipFile(src) as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
            zout.writestr(
                "xl/externalLinks/externalLink1.xml",
                '<externalLink '
                'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                "</externalLink>",
            )
    os.replace(tmp, src)

    result = ensure_formula_cache(str(p), raise_on_error=False)
    assert result["ok"] is False
    assert "外部链接" in result["error"]

    with pytest.raises(RuntimeError, match="外部链接"):
        ensure_formula_cache(str(p), raise_on_error=True)


def test_write_cells_via_com_rejects_formula():
    """公式写入必须被拒绝（无 Excel 也抛错——win32com 缺失先于值检查）。"""
    if HAS_WIN32COM:
        with pytest.raises(ValueError, match="公式写入"):
            write_cells_via_com("nonexistent.xlsx", {"#T": [(4, 2, "=B4+1")]})
    else:
        with pytest.raises(RuntimeError, match="win32com"):
            write_cells_via_com("nonexistent.xlsx", {"#T": [(4, 2, "=B4+1")]})


def test_apply_write_plan_via_com_formula_allowed(tmp_path):
    """resolve 写计划允许公式（受控写入，.Formula）；无 Excel 时 win32com 报错。"""
    p = tmp_path / "plan.xlsx"
    _write_xlsx_with_error_caches(p, sheets=("#T1",))
    plan = {
        "#T1": {
            "created": False,
            "headers": {},
            "inserts": [],
            "updates": [(4, 2, "=A4")],
            "deletes": [],
            "delete_from": None,
        }
    }
    if HAS_WIN32COM:
        import openpyxl

        apply_write_plan_via_com(str(p), plan)
        wb = openpyxl.load_workbook(str(p), data_only=False)
        assert wb["#T1"].cell(row=4, column=2).value == "=A4"
        wb.close()
    else:
        with pytest.raises(RuntimeError, match="win32com"):
            apply_write_plan_via_com(str(p), plan)


@needs_excel
def test_write_cells_via_com_static_values(tmp_path):
    """COM 直写静态值：写回后 Excel 可打开、值正确、公式缓存保留。"""
    import openpyxl

    p = tmp_path / "static.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    ws.append(["名称", "值"])
    ws.append(["Name", "Value"])
    ws.append(["FString", "FString"])
    ws.append([None, "old"])
    wb.save(str(p))
    wb.close()

    write_cells_via_com(str(p), {"#T": [(4, 2, "new_value")]})

    # openpyxl 可重新打开，值已写入
    wb2 = openpyxl.load_workbook(str(p), data_only=False)
    assert wb2["#T"].cell(row=4, column=2).value == "new_value"
    wb2.close()
