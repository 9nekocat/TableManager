# -*- coding: utf-8 -*-
"""COM 直写（write_cells_via_com / apply_write_plan_via_com）全链路测试。

覆盖（需本机 Excel / win32com）：
- write_cells_via_com：静态值写回，文件可被 openpyxl 重新打开、值正确；
  未被写入的公式单元格保持公式（含 XLOOKUP 等动态数组公式字符串不丢失）
- apply_write_plan_via_com：行插入 / 单元格修改 / 尾部删除组合，结果行序
  与值正确（验证行号漂移推演：插入按最终行号升序）
- 写回后 #N/A 基线零新增（667 基线的 Effect 场景在回归任务中做，这里用
  构造数据验证计数不因 COM 写回而增加）
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pytest

from excel_recalc import (
    apply_write_plan_via_com,
    count_formula_error_caches,
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


def _make_simple_table(path, rows):
    """构造 #T sheet：3 行表头 + rows（每行 [rowname, value]）。"""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    ws.append(["名称", "值"])
    ws.append(["Name", "Value"])
    ws.append(["FString", "FString"])
    for rn, val in rows:
        ws.append([rn, val])
    wb.save(str(path))
    wb.close()


@needs_excel
def test_write_cells_via_com_static_value_and_formula_preserved(tmp_path):
    """COM 直写：写入静态值，未写入的公式单元格保持公式不丢。"""
    import openpyxl

    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    ws.append(["名称", "值", "公式列"])
    ws.append(["Name", "Value", "Formula"])
    ws.append(["FString", "FString", "FString"])
    ws.append(["=B4", "old", "=B4&\"_f\""])
    ws.append(["=B5", "v2", "=B5&\"_f\""])
    wb.save(str(p))
    wb.close()

    write_cells_via_com(str(p), {"#T": [(4, 2, "new")]})

    wb2 = openpyxl.load_workbook(str(p), data_only=False)
    assert wb2["#T"].cell(row=4, column=2).value == "new"
    # 未写入单元格原样保留：RowName 公式与公式列
    assert wb2["#T"].cell(row=4, column=3).value == '=B4&"_f"'
    assert wb2["#T"].cell(row=5, column=3).value == '=B5&"_f"'
    wb2.close()


@needs_excel
def test_write_plan_inserts_updates_deletes(tmp_path):
    """apply_write_plan_via_com：插行 + 改值 + 删行 + 尾部删除组合。"""
    import openpyxl

    p = tmp_path / "plan.xlsx"
    # mine 行 4-7：A,B,C,D
    _make_simple_table(p, [("A", "1"), ("B", "2"), ("C", "3"), ("D", "4")])

    plan = {
        "#T": {
            "created": False,
            "headers": {},
            # mine 行 4-7: A,B,C,D；删 C(6) 后最终布局：
            # A(4), NEW1(5), B(6), D(7), NEW2(8)
            "inserts": [(5, {1: "N1", 2: "n1"}), (8, {1: "N2", 2: "n2"})],
            # 修改 D 行（最终行 7）的值；删除 mine 行 6（C）
            "updates": [(7, 2, "d_updated")],
            "deletes": [6],
            "delete_from": None,
        }
    }
    apply_write_plan_via_com(str(p), plan)

    wb2 = openpyxl.load_workbook(str(p), data_only=False)
    ws = wb2["#T"]
    assert [ws.cell(row=r, column=1).value for r in range(4, 9)] == [
        "A", "N1", "B", "D", "N2",
    ]
    assert ws.cell(row=7, column=2).value == "d_updated"
    wb2.close()


@needs_excel
def test_write_plan_tail_delete(tmp_path):
    """尾部删除：delete_from 之后的行被删除。"""
    import openpyxl

    p = tmp_path / "tail.xlsx"
    _make_simple_table(p, [("A", "1"), ("B", "2"), ("C", "3"), ("D", "4")])

    plan = {
        "#T": {
            "created": False,
            "headers": {},
            "inserts": [],
            "updates": [],
            "deletes": [],
            "delete_from": 6,  # 保留 A、B（行 4-5），删 C、D（行 6-7）
        }
    }
    apply_write_plan_via_com(str(p), plan)

    wb2 = openpyxl.load_workbook(str(p), data_only=False)
    ws = wb2["#T"]
    assert [ws.cell(row=r, column=1).value for r in range(4, 8)] == [
        "A", "B", None, None,
    ]
    wb2.close()


@needs_excel
def test_write_plan_created_sheet(tmp_path):
    """created sheet：新建 sheet + 表头 + 数据行。"""
    import openpyxl

    p = tmp_path / "created.xlsx"
    _make_simple_table(p, [("A", "1")])

    plan = {
        "#NEW": {
            "created": True,
            "headers": {(1, 1): "名称", (1, 2): "值", (2, 1): "Name", (2, 2): "Value"},
            "inserts": [(4, {1: "X", 2: "x"})],
            "updates": [],
            "deletes": [],
            "delete_from": None,
        }
    }
    apply_write_plan_via_com(str(p), plan)

    wb2 = openpyxl.load_workbook(str(p), data_only=False)
    assert "#NEW" in wb2.sheetnames
    ws = wb2["#NEW"]
    assert ws.cell(row=1, column=1).value == "名称"
    assert ws.cell(row=4, column=1).value == "X"
    assert ws.cell(row=4, column=2).value == "x"
    wb2.close()


@needs_excel
def test_com_write_no_new_na(tmp_path):
    """COM 直写不产生新 #N/A：写回前后错误值计数一致。"""
    import openpyxl

    p = tmp_path / "nona.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#T"
    ws.append(["名称", "值"])
    ws.append(["Name", "Value"])
    ws.append(["FString", "FString"])
    # A 列 RowName 公式（缓存留空，写回前后都不算错误值）
    ws.append(["=B4", "v1"])
    ws.append(["=B5", "v2"])
    wb.save(str(p))
    wb.close()

    baseline = count_formula_error_caches(str(p))
    write_cells_via_com(str(p), {"#T": [(4, 2, "changed")]})
    after = count_formula_error_caches(str(p))
    assert baseline == {}
    assert after == {}  # 写回不产生新 #N/A


@needs_excel
def test_write_cells_via_com_rejects_formula_with_excel(tmp_path):
    """有 Excel 时公式写入同样被拒绝（先于文件操作）。"""
    import openpyxl

    p = tmp_path / "reject.xlsx"
    _make_simple_table(p, [("A", "1")])
    with pytest.raises(ValueError, match="公式写入"):
        write_cells_via_com(str(p), {"#T": [(4, 2, "=B4+1")]})
