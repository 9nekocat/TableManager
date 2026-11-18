# -*- coding: utf-8 -*-
"""Add-merge（增量合并，staging -> target）写回行为测试。

覆盖现状实现：openpyxl 写回 + ArrayFormula 上下文包装 + zip 层恢复外链/
动态数组元数据 + Excel 重算公式缓存。重点回归 Card.xlsx 数组公式缓存损坏
（'1,10,24' -> '1'）的同类缺陷：修改/新增单元格时数组公式必须保持
t="array" 上下文，XLDAPR metadata 必须保留。
"""

import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from addmerge_engine import AddMergeError, apply_additive_merge
from merge_engine import _inject_dynamic_array_metadata

try:
    import win32com.client  # noqa: F401
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

needs_excel = pytest.mark.skipif(
    not HAS_WIN32COM, reason="requires Excel COM (win32com)"
)

# 与 Excel 保存的 Card.xlsx xl/metadata.xml 一致（XLDAPR 动态数组元数据）
XLDAPR_METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
    'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
    'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
    '</metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    '</ext></extLst></bk></futureMetadata>'
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
).encode("utf-8")

HEADER = [["名称", "值", "数组", "Extra"],
          ["Name", "Value", "Arr", "Extra2"],
          ["FString", "int32", "FString", "FString"]]


def _make_target(path, nrows=3, with_array=False, array_cell="D4"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    for row in HEADER:
        ws.append(row)
    for i in range(1, nrows + 1):
        ws.append([f"=B{i + 3}", f"row_{i}", f"v{i}", f"x{i}"])
    if with_array:
        ws[array_cell] = ArrayFormula(array_cell, "=SUM(C4:C6)")
    wb.save(path)
    wb.close()


def _make_staging(path, rows=None, array_formula=None):
    """rows: list of (name, value, arr, extra)；默认单行修改 row_1。

    array_formula: 形如 ("D4", "=SUM(C4:D4)") 时，把该格写成数组公式
    （模拟 staging 表自带数组公式列）。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    for row in HEADER:
        ws.append(row)
    for name, value, arr, extra in rows or [("row_1", "changed", "v1", "x1")]:
        ws.append([name, value, arr, extra])
    if array_formula:
        ws[array_formula[0]] = ArrayFormula(array_formula[0], array_formula[1])
    wb.save(path)
    wb.close()


def _inject_xldapr_metadata(target_path, tmp_path):
    """给 openpyxl 写的 xlsx 注入完整 XLDAPR metadata.xml + cm 属性。"""
    src = tmp_path / "metadata_src.zip"
    with zipfile.ZipFile(src, "w") as z:
        z.writestr("xl/metadata.xml", XLDAPR_METADATA)
    _inject_dynamic_array_metadata(str(src), str(target_path))


@pytest.fixture
def backup_dir(tmp_path, monkeypatch):
    import merge_workflow
    d = tmp_path / "backups"
    monkeypatch.setattr(merge_workflow, "BACKUP_DIR", d)
    return d


# ---------------------------------------------------------------------------
# 报告与校验
# ---------------------------------------------------------------------------

def test_dry_run_does_not_write(tmp_path):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    before = t.read_bytes()
    _make_staging(s)
    report = apply_additive_merge(str(s), str(t), apply=False)
    assert report["written"] is False
    assert report["sheets"]["#T"]["modified"], "应有修改行"
    assert t.read_bytes() == before


def test_apply_reports_written_and_backup(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s)
    report = apply_additive_merge(str(s), str(t), apply=True)
    assert report["written"] is True
    assert report["backup_path"] and Path(report["backup_path"]).exists()


def test_missing_files_raise(tmp_path):
    s = tmp_path / "staging.xlsx"
    _make_staging(s)
    with pytest.raises(AddMergeError, match="暂存表不存在"):
        apply_additive_merge(str(tmp_path / "nope.xlsx"), str(tmp_path / "t.xlsx"))
    with pytest.raises(AddMergeError, match="目标表不存在"):
        apply_additive_merge(str(s), str(tmp_path / "nope.xlsx"))


def test_unknown_sheet_skipped(tmp_path):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s)
    wb = openpyxl.load_workbook(s)
    ws = wb.active
    ws.title = "#Other"
    wb.save(s)
    wb.close()
    report = apply_additive_merge(str(s), str(t))
    sheet = report["sheets"]["#Other"]
    assert sheet["status"] == "skipped"
    assert sheet["skip_reason"]


def test_staging_duplicate_rownames_rejected(tmp_path):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s, rows=[("row_1", "a", "b", "c"), ("row_1", "a2", "b2", "c2")])
    with pytest.raises(AddMergeError, match="重复 RowName"):
        apply_additive_merge(str(s), str(t), apply=True)


# ---------------------------------------------------------------------------
# 写回行为（openpyxl 路径）
# ---------------------------------------------------------------------------

def test_apply_modified_cell(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s, rows=[("row_1", "changed", "v1", "")])
    report = apply_additive_merge(str(s), str(t), apply=True)
    assert report["written"] is True
    wb = openpyxl.load_workbook(t, data_only=False)
    assert wb["#T"]["B4"].value == "changed"
    wb.close()


def test_apply_appends_new_row(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s, rows=[("row_new", "nv", "nva", "nx")])
    report = apply_additive_merge(str(s), str(t), apply=True)
    assert report["written"] is True
    wb = openpyxl.load_workbook(t, data_only=False)
    assert wb["#T"]["A7"].value == "row_new"
    assert wb["#T"]["B7"].value == "nv"
    wb.close()


# ---------------------------------------------------------------------------
# 数组公式上下文保留（回归：Card.xlsx '1,10,24' -> '1'）
# ---------------------------------------------------------------------------

def test_array_formula_kept_on_modified_cell(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t, with_array=True)  # D4 = 数组公式 =SUM(C4:C6)
    # staging 改了 D4 的公式文本 → 触发覆盖；目标单元格是数组公式。
    # COM 直写路径拒绝公式写入（formula-write-reject spec）；降级路径
    # （无 Excel）保持数组公式上下文，否则重算按隐式交集取首值
    _make_staging(s, rows=[("row_1", "changed", "v1", "=SUM(C4:E4)")])
    if HAS_WIN32COM:
        with pytest.raises(AddMergeError, match="公式写入"):
            apply_additive_merge(str(s), str(t), apply=True)
        return
    apply_additive_merge(str(s), str(t), apply=True)
    wb = openpyxl.load_workbook(t, data_only=False)
    v = wb["#T"]["D4"].value
    assert isinstance(v, ArrayFormula), f"数组公式上下文丢失: {v!r}"
    assert v.text.startswith("=SUM"), f"被展平: {v!r}"
    wb.close()


def test_array_formula_kept_on_added_row(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    # staging 新增行（row_new）自带数组公式列 D4；写回目标追加行后
    # 必须保持数组上下文且行引用随目标行号重映射（降级路径行为；
    # COM 直写路径拒绝公式写入）
    _make_staging(s, rows=[("row_new", "nv", "nva", None)],
                  array_formula=("D4", "=SUM(C4:D4)"))
    if HAS_WIN32COM:
        with pytest.raises(AddMergeError, match="公式写入"):
            apply_additive_merge(str(s), str(t), apply=True)
        return
    apply_additive_merge(str(s), str(t), apply=True)
    wb = openpyxl.load_workbook(t, data_only=False)
    v = wb["#T"]["D7"].value
    assert isinstance(v, ArrayFormula), f"追加行数组公式丢失: {v!r}"
    assert v.text == "=SUM(C7:D7)", f"行引用未重映射: {v!r}"
    wb.close()


def test_xldapr_metadata_kept_after_apply(tmp_path, backup_dir):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t, with_array=True)
    _inject_xldapr_metadata(t, tmp_path)
    _make_staging(s, rows=[("row_1", "changed", "v1", "")])
    apply_additive_merge(str(s), str(t), apply=True)
    with zipfile.ZipFile(t) as z:
        assert "xl/metadata.xml" in z.namelist(), "XLDAPR metadata 丢失"
        assert b"XLDAPR" in z.read("xl/metadata.xml")
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert 'cm="1"' in sheet_xml, "数组公式单元格 cm 属性丢失"


# ---------------------------------------------------------------------------
# 公式缓存
# ---------------------------------------------------------------------------

@needs_excel
def test_apply_refreshes_caches_and_keeps_rowname_formula(tmp_path, backup_dir):
    import excel_recalc
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    # 先用 Excel 烘焙真实公式缓存（模拟现实工作副本；无缓存时 A 列公式
    # 会被 add-merge 误判为差异而替换成字面值——既有行为，非本测试目标）
    excel_recalc.recalc_xlsx(str(t), raise_on_error=True)
    assert excel_recalc.count_missing_formula_caches(str(t)) == 0

    _make_staging(s, rows=[("row_1", "changed", "v1", "")])
    report = apply_additive_merge(str(s), str(t), apply=True)
    assert report["written"] is True
    assert report["na_guard"].get("ok") is True

    wbf = openpyxl.load_workbook(t, data_only=False)
    assert wbf["#T"]["A4"].value == "=B4"  # RowName 公式保留
    wbf.close()
    wb = openpyxl.load_workbook(t, data_only=True)
    assert wb["#T"]["B4"].value == "changed"
    assert wb["#T"]["A4"].value == "changed"  # =B4 依赖被重算
    wb.close()


def test_recalc_degrades_without_win32com(tmp_path, backup_dir, monkeypatch):
    t = tmp_path / "target.xlsx"
    s = tmp_path / "staging.xlsx"
    _make_target(t)
    _make_staging(s, rows=[("row_1", "changed", "v1", "")])
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    report = apply_additive_merge(str(s), str(t), apply=True)
    assert report["written"] is True
    assert report["na_guard"].get("ok") is None
    assert "skipped" in report["na_guard"]
    # 写回本身不依赖 Excel，单元格仍被修改
    wb = openpyxl.load_workbook(t, data_only=False)
    assert wb["#T"]["B4"].value == "changed"
    wb.close()
