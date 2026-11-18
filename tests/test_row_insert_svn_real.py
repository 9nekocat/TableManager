# -*- coding: utf-8 -*-
"""Test Bug 2 with real SVN artifact filenames (non-standard extensions)."""
import sys
import os
import tempfile
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
from merge_engine import generate_merge_preview, resolve_merge

DATA_START_ROW = 4


def _make_wb(path, rows_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#TestSheet"
    for i, h in enumerate(["RowName", "Name", "Value"], 1):
        ws.cell(row=1, column=i, value=h)
        ws.cell(row=2, column=i, value=h)
        ws.cell(row=3, column=i, value="FString")
    for ri, row_data in enumerate(rows_data, DATA_START_ROW):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val)
    wb.save(path)
    wb.close()


def test_svn_artifact_filenames():
    """
    Simulate real SVN conflict: Effect.xlsx has conflicts, SVN produces
    Effect.xlsx.mine (mine) and Effect.xlsx.r4524 (theirs).

    When theirs_path has .r4524 extension, resolve_merge must handle it
    and correctly position THEIRS_ADDED rows.
    """
    tmp = tempfile.mkdtemp()
    try:
        # Use SVN-like filenames
        base_path = os.path.join(tmp, "Effect_base.xlsx")
        mine_path = os.path.join(tmp, "Effect.xlsx.mine")    # SVN .mine extension
        theirs_path = os.path.join(tmp, "Effect.xlsx.r4524")   # SVN r-prefix extension
        preview_path = os.path.join(tmp, "preview.xlsx")
        resolved_path = os.path.join(tmp, "Effect.xlsx")

        _make_wb(base_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("E", "Echo", "5"),
        ])
        _make_wb(mine_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("E", "Echo", "5"),
        ])
        _make_wb(theirs_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("C", "Charlie", "3"),
            ("D", "Delta", "4"),
            ("E", "Echo", "5"),
        ])

        # Preview: need to handle the extension issue
        preview_result = generate_merge_preview(
            base_path, mine_path, theirs_path, preview_path
        )
        print(f"Preview generated: {preview_result}")

        # Resolve with SVN-like theirs_path
        result = resolve_merge(preview_path, resolved_path,
                                mine_path=mine_path, theirs_path=theirs_path)

        # Read resolved
        wb = openpyxl.load_workbook(result, data_only=True)
        ws = wb["#TestSheet"]
        resolved_names = []
        for row_idx in range(DATA_START_ROW, ws.max_row + 1):
            name = ws.cell(row=row_idx, column=2).value
            if name:
                resolved_names.append(str(name).strip())
        wb.close()

        expected = ["Alpha", "Bravo", "Charlie", "Delta", "Echo"]
        if resolved_names == expected:
            print(f"PASS: SVN artifact filenames handled correctly")
            print(f"  Order: {resolved_names}")
            return True
        else:
            print(f"FAIL: Wrong row order with SVN artifact filenames")
            print(f"  Expected: {expected}")
            print(f"  Got:      {resolved_names}")
            return False

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    ok = test_svn_artifact_filenames()
    sys.exit(0 if ok else 1)
