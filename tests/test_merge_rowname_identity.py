# -*- coding: utf-8 -*-
"""Regression: theirs modifying a row's B column (part of the (colA, colB)
preview key) must resolve as an in-place modification, not as an inserted
new row — inserting duplicates the RowName (real case: GameplayTagTableRow
#Stage 序号4, 驻留商店/Stage.Shop.Sell → 特殊商店前置事件/Stage.Event.Router).
"""
from collections import Counter

import openpyxl

from merge_engine import generate_merge_preview, resolve_merge

_HEADERS = [
    ("序号", "注释", "标签"),
    ("Row", "DevComment", "Tag"),
    ("int", "string", "string"),
]

_BASE_ROWS = [
    (1, "Flag体系", "Flag"),
    (2, "世界状态", "Flag.World"),
    (3, "对话状态", "Flag.Dialog"),
    (4, "驻留商店", "Stage.Shop.Sell"),
    (5, "随机商店", "Stage.Shop.Random"),
]


def _write_book(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#Stage"
    for r_idx, row in enumerate(_HEADERS, start=1):
        for c_idx, cell_value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=cell_value)
    for r_idx, row in enumerate(rows, start=4):
        for c_idx, cell_value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=cell_value)
    wb.save(path)
    wb.close()


def _data_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb["#Stage"]
        return [tuple(r[:3]) for r in ws.iter_rows(min_row=4, values_only=True)
                if r[0] is not None]
    finally:
        wb.close()


def test_theirs_modifying_col_b_does_not_duplicate_rowname(tmp_path):
    base = tmp_path / "base.xlsx"
    mine = tmp_path / "mine.xlsx"
    theirs = tmp_path / "theirs.xlsx"
    preview = tmp_path / "preview_merge.xlsx"
    resolved = tmp_path / "resolved.xlsx"

    _write_book(base, _BASE_ROWS)
    _write_book(mine, _BASE_ROWS)  # mine keeps the row unchanged
    theirs_rows = [
        row if row[0] != 4 else (4, "特殊商店前置事件", "Stage.Event.Router")
        for row in _BASE_ROWS
    ]
    _write_book(theirs, theirs_rows)

    generate_merge_preview(str(base), str(mine), str(theirs), str(preview))
    resolve_merge(str(preview), str(resolved),
                  mine_path=str(mine), theirs_path=str(theirs), base_path=str(base))

    rows = _data_rows(resolved)
    counts = Counter(r[0] for r in rows)
    assert counts[4] == 1, f"RowName 4 duplicated: {rows}"
    row4 = next(r for r in rows if r[0] == 4)
    assert row4 == (4, "特殊商店前置事件", "Stage.Event.Router")
    # No other row lost or duplicated
    assert [r[0] for r in rows] == [1, 2, 3, 4, 5]
