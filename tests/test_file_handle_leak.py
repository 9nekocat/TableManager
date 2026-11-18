# -*- coding: utf-8 -*-
"""Test for Bug 1: File handle leak after merge (blocks external SVN operations)."""
import sys
import os
import tempfile
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import openpyxl
from merge_engine import generate_merge_preview, resolve_merge


def _make_test_files(tmpdir):
    """Create minimal base/mine/theirs files for merge test."""
    files = {}
    for role in ('base', 'mine', 'theirs'):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "#TestSheet"
        for i, h in enumerate(["RowName", "Name", "Value"], 1):
            ws.cell(row=1, column=i, value=h)
            ws.cell(row=2, column=i, value=h)
            ws.cell(row=3, column=i, value="FString")

        data = {
            'base':   [("A", "Alice", "100"), ("B", "Bob", "200")],
            'mine':   [("A", "Alice_mod", "100"), ("C", "Carol", "300")],
            'theirs': [("A", "Alice", "150"), ("D", "Dave", "400")],
        }
        for ri, row_data in enumerate(data[role], 4):
            for ci, val in enumerate(row_data, 1):
                ws.cell(row=ri, column=ci, value=val)

        path = os.path.join(tmpdir, f"{role}.xlsx")
        wb.save(path)
        wb.close()
        files[role] = path
    return files


def test_handles_closed_after_merge():
    """After resolve_merge completes, all input files must be deletable."""
    tmp = tempfile.mkdtemp()
    try:
        files = _make_test_files(tmp)
        preview_path = os.path.join(tmp, "preview.xlsx")
        resolved_path = os.path.join(tmp, "resolved.xlsx")

        # Step 1: Generate preview
        generate_merge_preview(
            files['base'], files['mine'], files['theirs'], preview_path
        )

        # Step 2: Resolve
        resolve_merge(preview_path, resolved_path,
                      mine_path=files['mine'],
                      theirs_path=files['theirs'])

        # Step 3: Try to delete all temp files — if handles are open, this fails
        for name, path in files.items():
            try:
                os.remove(path)
            except PermissionError as e:
                print(f"FAIL: Cannot delete {name} after merge: {e}")
                return False

        # Also clean up preview and resolved
        for p in (preview_path, resolved_path):
            if os.path.exists(p):
                os.remove(p)

        print("PASS: All file handles released after merge")
        return True
    finally:
        try:
            shutil.rmtree(tmp, ignore_errors=True)
        except Exception:
            pass  # If files are locked, rmtree will fail silently


if __name__ == "__main__":
    ok = test_handles_closed_after_merge()
    sys.exit(0 if ok else 1)
