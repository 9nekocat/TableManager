# -*- coding: utf-8 -*-
"""Test for Bug 2: THEIRS_ADDED rows inserted at end instead of correct position."""
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
    """Create minimal workbook with rows_data = list of (RowName, Name, Value)."""
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


def test_theirs_added_row_position():
    """
    THEIRS_ADDED row should insert between existing rows, not at end.

    Scenario:
      base:   A, B, C, E
      mine:   A, B, C, E   (same as base)
      theirs: A, B, C, D, E  (inserted D between C and E)

    Expected: resolved order = A, B, C, D, E
    Actual (bug): resolved order = A, B, C, E, D (D at end)
    """
    tmp = tempfile.mkdtemp()
    try:
        base_path = os.path.join(tmp, "base.xlsx")
        mine_path = os.path.join(tmp, "mine.xlsx")
        theirs_path = os.path.join(tmp, "theirs.xlsx")
        preview_path = os.path.join(tmp, "preview.xlsx")
        resolved_path = os.path.join(tmp, "resolved.xlsx")

        _make_wb(base_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("C", "Charlie", "3"),
            ("E", "Echo", "5"),
        ])
        _make_wb(mine_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("C", "Charlie", "3"),
            ("E", "Echo", "5"),
        ])
        _make_wb(theirs_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
            ("C", "Charlie", "3"),
            ("D", "Delta", "4"),    # NEW — between C and E
            ("E", "Echo", "5"),
        ])

        # Generate preview + resolve with correct theirs_path
        generate_merge_preview(base_path, mine_path, theirs_path, preview_path)
        resolve_merge(preview_path, resolved_path,
                      mine_path=mine_path, theirs_path=theirs_path)

        # Read resolved data rows and check order
        wb = openpyxl.load_workbook(resolved_path, data_only=True)
        ws = wb["#TestSheet"]
        resolved_names = []
        for row_idx in range(DATA_START_ROW, ws.max_row + 1):
            name = ws.cell(row=row_idx, column=2).value
            if name:
                resolved_names.append(str(name).strip())
        wb.close()

        expected_order = ["Alpha", "Bravo", "Charlie", "Delta", "Echo"]
        if resolved_names == expected_order:
            print("PASS: THEIRS_ADDED row inserted at correct position")
            print(f"  Order: {resolved_names}")
            return True
        else:
            print(f"FAIL: Wrong row order after merge")
            print(f"  Expected: {expected_order}")
            print(f"  Got:      {resolved_names}")
            return False

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_theirs_added_missing_path_fallback():
    """
    When theirs_path is NOT passed (None), row should still insert
    at best-effort position, not crash.
    """
    tmp = tempfile.mkdtemp()
    try:
        base_path = os.path.join(tmp, "base.xlsx")
        mine_path = os.path.join(tmp, "mine.xlsx")
        theirs_path = os.path.join(tmp, "theirs.xlsx")
        preview_path = os.path.join(tmp, "preview.xlsx")
        resolved_path = os.path.join(tmp, "resolved.xlsx")

        _make_wb(base_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
        ])
        _make_wb(mine_path, [
            ("A", "Alpha", "1"),
            ("B", "Bravo", "2"),
        ])
        _make_wb(theirs_path, [
            ("A", "Alpha", "1"),
            ("C", "Charlie", "3"),    # NEW — between A and B? At end?
            ("B", "Bravo", "2"),
        ])

        generate_merge_preview(base_path, mine_path, theirs_path, preview_path)

        # Resolve WITHOUT theirs_path — simulates edge case
        result_path = resolve_merge(preview_path, resolved_path,
                                     mine_path=mine_path, theirs_path=None)

        wb = openpyxl.load_workbook(result_path, data_only=True)
        ws = wb["#TestSheet"]
        resolved_names = []
        for row_idx in range(DATA_START_ROW, ws.max_row + 1):
            name = ws.cell(row=row_idx, column=2).value
            if name:
                resolved_names.append(str(name).strip())
        wb.close()

        # With theirs_path=None, the row may end up at end. Check we don't crash.
        if len(resolved_names) == 3:
            print(f"PASS: Resolve succeeds without theirs_path (order may vary)")
            print(f"  Resolved order: {resolved_names}")
            return True
        else:
            print(f"FAIL: Wrong row count: got {len(resolved_names)}, expected 3")
            return False

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    r1 = test_theirs_added_row_position()
    print()
    r2 = test_theirs_added_missing_path_fallback()
    sys.exit(0 if (r1 and r2) else 1)
