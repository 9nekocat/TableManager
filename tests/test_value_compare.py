# -*- coding: utf-8 -*-
"""Position-based value comparison: generated resolved vs _resolved_correct.
Compares cells row-by-row in order, avoiding RowName resolution issues."""

import sys
from pathlib import Path
from collections import OrderedDict

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from manager_core import DATA_START_ROW

TEST_DIR = Path(__file__).resolve().parent.parent / "测试用例"
OUTPUT_DIR = TEST_DIR / "_outputs"


def _read_rows_ordered(filepath):
    """Read all #-sheet data rows in order. Returns {sheet: [(row_values, row_formulas)]}.
    Each row: values = tuple of cell values (data_only), formulas = tuple of cell formulas."""
    import openpyxl
    import tempfile, shutil

    actual = filepath
    tmp = None
    if not str(filepath).lower().endswith('.xlsx'):
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        shutil.copy2(filepath, tmp.name)
        actual = tmp.name

    try:
        wb_val = openpyxl.load_workbook(actual, data_only=True, read_only=True)
        wb_fml = openpyxl.load_workbook(actual, data_only=False, read_only=True)
    except Exception as e:
        raise IOError(f"Cannot open: {e}")
    finally:
        if tmp:
            import os
            try: os.unlink(tmp.name)
            except: pass

    result = {}
    for sn in wb_val.sheetnames:
        if not sn.startswith("#"):
            continue
        ws_val = wb_val[sn]
        val_rows = list(ws_val.iter_rows(min_row=DATA_START_ROW, values_only=True))

    for sn in wb_fml.sheetnames:
        if not sn.startswith("#"):
            continue
        ws_fml = wb_fml[sn]
        fml_rows = list(ws_fml.iter_rows(min_row=DATA_START_ROW, values_only=True))

    for sn in wb_val.sheetnames:
        if not sn.startswith("#"):
            continue
        ws_val = wb_val[sn]
        val_rows = list(ws_val.iter_rows(min_row=DATA_START_ROW, values_only=True))
        fml_rows_raw = list(wb_fml[sn].iter_rows(min_row=DATA_START_ROW, values_only=True)) if sn in wb_fml.sheetnames else []

        rows = []
        for i, vrow in enumerate(val_rows):
            frow = fml_rows_raw[i] if i < len(fml_rows_raw) else tuple()
            if not vrow:
                continue
            vals = tuple(str(v).strip() if v is not None else "" for v in vrow)
            fmls = tuple(str(f).strip() if isinstance(f, str) and str(f).startswith("=") else "" for f in frow)
            rows.append((vals, fmls))
        result[sn] = rows

    wb_val.close()
    wb_fml.close()
    return result


def compare(correct_path, generated_path):
    """Position-based cell comparison. Returns (pass, diffs_list)."""
    correct = _read_rows_ordered(correct_path)
    generated = _read_rows_ordered(generated_path)

    diffs = []
    all_sheets = sorted(set(list(correct.keys()) + list(generated.keys())))

    for sn in all_sheets:
        c_rows = correct.get(sn, [])
        g_rows = generated.get(sn, [])

        if len(c_rows) != len(g_rows):
            diffs.append(f"[{sn}] Row count: correct={len(c_rows)}, generated={len(g_rows)}")

        max_rows = max(len(c_rows), len(g_rows))
        cell_diffs = 0
        for pos in range(max_rows):
            if pos >= len(c_rows):
                diffs.append(f"[{sn}] Row pos {pos+1}: only in generated")
                continue
            if pos >= len(g_rows):
                diffs.append(f"[{sn}] Row pos {pos+1}: only in correct")
                continue

            cv, cf = c_rows[pos]
            gv, gf = g_rows[pos]
            max_cols = max(len(cv), len(gv), len(cf), len(gf))

            row_diffs = []
            for col in range(max_cols):
                c_val = cv[col] if col < len(cv) else ""
                g_val = gv[col] if col < len(gv) else ""
                c_fml = cf[col] if col < len(cf) else ""
                g_fml = gf[col] if col < len(gf) else ""

                # Compare: prefer formula if both have, else value
                if c_fml and g_fml:
                    if c_fml != g_fml:
                        row_diffs.append(f"  col{col} fml: [{c_fml[:80]}] vs [{g_fml[:80]}]")
                elif c_fml:
                    if c_fml.strip("=") != g_val:
                        row_diffs.append(f"  col{col}: correct fml [{c_fml[:60]}] vs gen val [{g_val[:60]}]")
                elif g_fml:
                    if c_val != g_fml.strip("="):
                        row_diffs.append(f"  col{col}: correct val [{c_val[:60]}] vs gen fml [{g_fml[:60]}]")
                else:
                    if c_val != g_val:
                        row_diffs.append(f"  col{col}: [{c_val[:50]}] vs [{g_val[:50]}]")

            if row_diffs:
                cell_diffs += 1
                if cell_diffs <= 5:
                    # Show RowName context for this row
                    rn_c = cv[0] if len(cv) > 0 else "?"
                    rn_g = gv[0] if len(gv) > 0 else "?"
                    diffs.append(f"[{sn}] Row pos {pos+1} ({rn_c}/{rn_g}):")
                    diffs.extend(row_diffs)

        if cell_diffs > 5:
            diffs.append(f"[{sn}] ... {cell_diffs - 5} more rows with differences")

    return (len(diffs) == 0), diffs


def main():
    print("=" * 60)
    print("Position-Based Cell Comparison")
    print("=" * 60)

    all_pass = True
    for correct_f in sorted(TEST_DIR.glob("*_resolved_correct.xlsx")):
        prefix = correct_f.stem.replace("_resolved_correct", "")
        generated_f = OUTPUT_DIR / f"{prefix}_resolved.xlsx"
        if not generated_f.exists():
            print(f"\n{prefix}: SKIP (no output)")
            continue

        ok, diffs = compare(str(correct_f), str(generated_f))
        if ok:
            print(f"\n{prefix}: PASS")
        else:
            all_pass = False
            print(f"\n{prefix}: FAIL — {len(diffs)} issue(s):")
            for d in diffs[:20]:
                print(f"  {d}")
            if len(diffs) > 20:
                print(f"  ... and {len(diffs)-20} more")

    print(f"\n{'='*60}")
    print("ALL PASS" if all_pass else "SOME FAILED")
    print("=" * 60)
    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
