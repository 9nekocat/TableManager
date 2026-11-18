# -*- coding: utf-8 -*-
"""Run all test cases from 测试用例/ and report errors."""
import os
import sys
import traceback
from pathlib import Path

os.chdir(Path(__file__).parent)
sys.path.insert(0, str(Path(__file__).parent))

from merge_engine import generate_merge_preview, resolve_merge
import uuid
import openpyxl

TEST_DIR = Path(__file__).resolve().parent.parent / "测试用例"
OUTPUT_DIR = TEST_DIR / "_output"
OUTPUT_DIR.mkdir(exist_ok=True)

# TortoiseSVN standard conflict naming:
#   TC_Name.xlsx.rOLDREV  = base (older revision)
#   TC_Name.xlsx.mine     = mine (working copy)
#   TC_Name.xlsx.rNEWREV  = theirs (incoming revision)
CASES = {
    "TC01_基础合并":    ("TC01_基础合并.xlsx.r1000", "TC01_基础合并.xlsx.mine", "TC01_基础合并.xlsx.r1001", "TC01_基础合并_resolved_correct.xlsx"),
    "TC02_多Sheet合并": ("TC02_多Sheet合并.xlsx.r1002", "TC02_多Sheet合并.xlsx.mine", "TC02_多Sheet合并.xlsx.r1003", "TC02_多Sheet合并_resolved_correct.xlsx"),
    "TC03_公式合并":    ("TC03_公式合并.xlsx.r1004", "TC03_公式合并.xlsx.mine", "TC03_公式合并.xlsx.r1005", "TC03_公式合并_resolved_correct.xlsx"),
    "TC04_增删合并":    ("TC04_增删合并.xlsx.r1006", "TC04_增删合并.xlsx.mine", "TC04_增删合并.xlsx.r1007", "TC04_增删合并_resolved_correct.xlsx"),
    "TC05_跨文件公式":  ("TC05_跨文件公式.xlsx.r1008", "TC05_跨文件公式.xlsx.mine", "TC05_跨文件公式.xlsx.r1009", "TC05_跨文件公式_resolved_correct.xlsx"),
    "TC06_无冲突":      ("TC06_无冲突.xlsx.r1010", "TC06_无冲突.xlsx.mine", "TC06_无冲突.xlsx.r1011", "TC06_无冲突_resolved_correct.xlsx"),
    "TC07_同名新增":    ("TC07_同名新增.xlsx.r1012", "TC07_同名新增.xlsx.mine", "TC07_同名新增.xlsx.r1013", "TC07_同名新增_resolved_correct.xlsx"),
    "TC08_特殊值":      ("TC08_特殊值.xlsx.r1014", "TC08_特殊值.xlsx.mine", "TC08_特殊值.xlsx.r1015", "TC08_特殊值_resolved_correct.xlsx"),
    "TC09_Sheet增删":   ("TC09_Sheet增删.xlsx.r1016", "TC09_Sheet增删.xlsx.mine", "TC09_Sheet增删.xlsx.r1017", "TC09_Sheet增删_resolved_correct.xlsx"),
    "TC10_无差异":      ("TC10_无差异.xlsx.r1018", "TC10_无差异.xlsx.mine", "TC10_无差异.xlsx.r1019", "TC10_无差异_resolved_correct.xlsx"),
    "TC11_大表":        ("TC11_大表.xlsx.r1020", "TC11_大表.xlsx.mine", "TC11_大表.xlsx.r1021", "TC11_大表_resolved_correct.xlsx"),
    "TC12_Row为公式":   ("TC12_Row为公式.xlsx.r1022", "TC12_Row为公式.xlsx.mine", "TC12_Row为公式.xlsx.r1023", "TC12_Row为公式_resolved_correct.xlsx"),
    "TC14_插入ROW为复杂公式": ("TC14_插入ROW为复杂公式.xlsx.r1024", "TC14_插入ROW为复杂公式.xlsx.mine", "TC14_插入ROW为复杂公式.xlsx.r1025", "TC14_插入ROW为复杂公式_resolved_correct.xlsx"),
    "TC15_综合": ("TC15_综合.xlsx.r1026", "TC15_综合.xlsx.mine", "TC15_综合.xlsx.r1027", "TC15_综合_correct.xlsx"),
    "TC16_列增删移动": ("TC16_列增删移动.xlsx.r1028", "TC16_列增删移动.xlsx.mine", "TC16_列增删移动.xlsx.r1029", "TC16_列增删移动_resolved_correct.xlsx"),
}

results = []

def compare_xlsx(resolved_path, correct_path):
    """Compare two xlsx files cell by cell using cached values.
    Both files must be pre-recalculated by Excel so formula cells have values."""
    diffs = []
    wb_r = openpyxl.load_workbook(resolved_path, data_only=True)
    wb_c = openpyxl.load_workbook(correct_path, data_only=True)

    all_sheets = set(wb_r.sheetnames) | set(wb_c.sheetnames)
    for sn in sorted(all_sheets):
        if sn not in wb_r.sheetnames:
            diffs.append(f"Sheet '{sn}' missing in resolved")
            continue
        if sn not in wb_c.sheetnames:
            diffs.append(f"Sheet '{sn}' extra in resolved")
            continue
        ws_r = wb_r[sn]
        ws_c = wb_c[sn]
        if ws_r.max_row != ws_c.max_row or ws_r.max_column != ws_c.max_column:
            diffs.append(f"Sheet '{sn}' size: resolved({ws_r.max_row}r x {ws_r.max_column}c) vs correct({ws_c.max_row}r x {ws_c.max_column}c)")
        for row_idx in range(1, max(ws_r.max_row, ws_c.max_row) + 1):
            for col_idx in range(1, max(ws_r.max_column, ws_c.max_column) + 1):
                r_val = ws_r.cell(row=row_idx, column=col_idx).value
                c_val = ws_c.cell(row=row_idx, column=col_idx).value
                r_s = "" if r_val is None else str(r_val)
                c_s = "" if c_val is None else str(c_val)
                if r_s != c_s:
                    col_letter = chr(64 + col_idx) if col_idx <= 26 else f"Col{col_idx}"
                    diffs.append(f"{sn}!{col_letter}{row_idx}: [{r_val}] vs [{c_val}]")

    wb_r.close()
    wb_c.close()
    return diffs


for case_name, (base_f, mine_f, theirs_f, correct_f) in CASES.items():
    base_p = str(TEST_DIR / base_f)
    mine_p = str(TEST_DIR / mine_f)
    theirs_p = str(TEST_DIR / theirs_f)
    correct_p = str(TEST_DIR / correct_f) if correct_f else None
    uid = uuid.uuid4().hex[:6]
    out_p = str(OUTPUT_DIR / f"{case_name}_merge_{uid}.xlsx")
    resolved_p = str(OUTPUT_DIR / f"{case_name}_resolved_{uid}.xlsx")

    print(f"\n{'='*60}")
    print(f"Running: {case_name}")
    print(f"  base:   {base_f}")
    print(f"  mine:   {mine_f}")
    print(f"  theirs: {theirs_f}")

    if not Path(base_p).exists():
        print(f"  [SKIP] base file missing: {base_p}")
        results.append((case_name, "SKIP", "base file missing"))
        continue

    # Step 0: Recalculate input files so formula cells have cached values
    # for value-based key matching in resolve_merge.
    try:
        from excel_recalc import recalc_xlsx
        recalc_xlsx(base_p, timeout=60)
        recalc_xlsx(mine_p, timeout=60)
        recalc_xlsx(theirs_p, timeout=60)
    except Exception:
        pass  # Excel not available; merge uses formula-string fallback

    # Step 1: Generate merge preview
    try:
        preview_path = generate_merge_preview(base_p, mine_p, theirs_p, out_p)
        if not Path(preview_path).exists():
            raise RuntimeError(f"Preview file not created: {preview_path}")
        print(f"  [OK]  Preview: {preview_path}")
    except Exception as e:
        print(f"  [FAIL] Preview error: {e}")
        traceback.print_exc()
        results.append((case_name, "FAIL_PREVIEW", str(e)))
        continue

    # Step 2: Resolve merge (pass theirs_path for correct row positioning)
    try:
        resolved = resolve_merge(preview_path, resolved_p, mine_path=mine_p, theirs_path=theirs_p)
        if not Path(str(resolved)).exists():
            raise RuntimeError(f"Resolved file not created: {resolved}")
        print(f"  [OK]  Resolved: {resolved}")
    except Exception as e:
        print(f"  [FAIL] Resolve error: {e}")
        traceback.print_exc()
        results.append((case_name, "FAIL_RESOLVE", str(e)))
        continue

    # Step 3: Compare with correct output
    if correct_p and Path(correct_p).exists():
        try:
            # Recalculate both files so formula cells have cached values
            try:
                from excel_recalc import recalc_xlsx
                recalc_xlsx(resolved, timeout=60)
                recalc_xlsx(correct_p, timeout=60)
            except Exception:
                pass

            diffs = compare_xlsx(resolved, correct_p)
            if diffs:
                print(f"  [FAIL] Correctness check: {len(diffs)} differences")
                for d in diffs[:10]:
                    print(f"         {d}")
                if len(diffs) > 10:
                    print(f"         ... and {len(diffs) - 10} more")
                results.append((case_name, "FAIL_CORRECT", f"{len(diffs)} diffs"))
            else:
                print(f"  [OK]  Correctness: all cells match")
                results.append((case_name, "PASS", "correct"))
        except Exception as e:
            print(f"  [FAIL] Correctness check error: {e}")
            results.append((case_name, "FAIL_CORRECT", str(e)))
    else:
        results.append((case_name, "PASS", f"resolved={Path(str(resolved)).name}"))

# Summary
print(f"\n{'='*60}")
print(f"{'TEST SUMMARY':^60}")
print(f"{'='*60}")

pass_count = sum(1 for r in results if r[1] == "PASS")
fail_count = sum(1 for r in results if r[1].startswith("FAIL"))
skip_count = sum(1 for r in results if r[1] == "SKIP")

for name, status, detail in results:
    icon = "[PASS]" if status == "PASS" else ("[FAIL]" if status.startswith("FAIL") else "[SKIP]")
    print(f"  {icon} {name}: {detail}")

print(f"\nTotal: {len(results)} | Pass: {pass_count} | Fail: {fail_count} | Skip: {skip_count}")

if fail_count > 0:
    sys.exit(1)
