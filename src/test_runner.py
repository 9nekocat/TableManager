# -*- coding: utf-8 -*-
"""Run all merge test cases under A:/Tools/TableManager/测试用例/."""

import os
import sys
import tempfile
import shutil
import hashlib
import zipfile
from pathlib import Path
from collections import OrderedDict

# Fix GBK encoding issues on Windows terminal
try:
    if sys.stdout.encoding and sys.stdout.encoding.lower() in ('gbk', 'cp936'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent))

from merge_engine import generate_merge_preview, resolve_merge

TEST_DIR = Path(__file__).resolve().parent.parent / "测试用例"
OUTPUT_DIR = TEST_DIR / "_outputs"


def find_test_cases():
    """Discover test cases with TortoiseSVN standard conflict naming.

    Patterns (stem = TC_Name without .xlsx suffix):
      TC_Name.xlsx.rREV  = base or theirs (lower rev = base, higher = theirs)
      TC_Name.xlsx.mine  = mine
    """
    import re
    cases = OrderedDict()

    for f in sorted(TEST_DIR.glob("*.xlsx*")):
        fname = f.name
        # Skip resolved_correct, correct, merge, temp files
        if any(kw in fname for kw in ("_resolved_correct", "_correct", "_merge", "~$")):
            continue

        # Match: TC_Name.xlsx.mine → stem = TC_Name
        if fname.endswith(".xlsx.mine"):
            stem = fname[:-10]  # strip .xlsx.mine
            cases.setdefault(stem, {})["mine"] = str(f)

        # Match: TC_Name.xlsx.rNNNN → stem = TC_Name
        m = re.match(r"(.+)\.xlsx\.r(\d+)$", fname)
        if m:
            stem = m.group(1)
            rev = int(m.group(2))
            entry = cases.setdefault(stem, {})
            entry.setdefault("rev_files", []).append((rev, str(f)))

    # Sort rev_files: lower rev = base, higher rev = theirs
    for prefix, entry in cases.items():
        rev_files = entry.get("rev_files", [])
        if rev_files:
            rev_files.sort(key=lambda x: x[0])
            if len(rev_files) >= 2:
                entry["base"] = rev_files[0][1]
                entry["theirs"] = rev_files[1][1]
            elif len(rev_files) == 1:
                entry["theirs"] = rev_files[0][1]
            del entry["rev_files"]

    return cases


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def compare_values_only(path_a, path_b):
    """Compare two xlsx files cell by cell using cached values.
    Both files must be pre-recalculated by Excel."""
    import openpyxl
    wb_a = openpyxl.load_workbook(path_a, data_only=True)
    wb_b = openpyxl.load_workbook(path_b, data_only=True)
    diffs = []
    all_sheets = set(wb_a.sheetnames) | set(wb_b.sheetnames)
    for sn in sorted(all_sheets):
        if sn not in wb_a.sheetnames:
            diffs.append(f"Sheet '{sn}' missing in A")
            continue
        if sn not in wb_b.sheetnames:
            diffs.append(f"Sheet '{sn}' extra in A")
            continue
        ws_a = wb_a[sn]
        ws_b = wb_b[sn]
        for r in range(1, max(ws_a.max_row, ws_b.max_row) + 1):
            for c in range(1, max(ws_a.max_column, ws_b.max_column) + 1):
                va = ws_a.cell(row=r, column=c).value
                vb = ws_b.cell(row=r, column=c).value
                if str(va) != str(vb):
                    diffs.append(f"{sn}!{chr(64+c) if c<=26 else f'Col{c}'}{r}: [{va}] vs [{vb}]")
    wb_a.close()
    wb_b.close()
    return (len(diffs) == 0), diffs


def run_one_test(prefix, roles):
    base = roles.get("base")
    mine = roles.get("mine")
    theirs = roles.get("theirs") or roles.get("their")

    if not all([base, mine, theirs]):
        return {"status": "SKIP", "reason": "missing base/mine/theirs"}

    # Check for expected output
    correct_file = None
    for f in TEST_DIR.glob(f"{prefix}_resolved_correct.xlsx"):
        correct_file = str(f)
        break

    work = tempfile.mkdtemp(prefix=f"{prefix}_")
    try:
        preview = os.path.join(work, f"{prefix}_merge.xlsx")
        resolved = os.path.join(work, f"{prefix}_resolved.xlsx")

        # Step 1: Generate merge preview
        preview_result = generate_merge_preview(base, mine, theirs, preview)
        if not os.path.isfile(preview_result):
            return {"status": "FAIL", "reason": "preview generation failed"}

        # Step 2: Resolve (pass theirs_path for correct row positioning)
        try:
            result_path = resolve_merge(preview, resolved, mine_path=mine, theirs_path=theirs)
            # NOTE: resolve_merge signature already accepts theirs_path;
            # we pass it here so THEIRS_ADDED rows insert at correct positions.
        except Exception as e:
            return {"status": "FAIL", "reason": f"resolve exception: {e}"}
        if not os.path.isfile(resolved):
            return {"status": "FAIL", "reason": f"resolve did not create output: {resolved}"}

        # Copy outputs to _outputs/
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        shutil.copy2(preview, OUTPUT_DIR / f"{prefix}_merge.xlsx")
        shutil.copy2(resolved, OUTPUT_DIR / f"{prefix}_resolved.xlsx")

        # Step 3: Verify against expected if available
        if correct_file and os.path.isfile(correct_file):
            ok, diffs = compare_values_only(resolved, correct_file)
            if ok:
                return {"status": "PASS", "reason": "matches _resolved_correct"}
            else:
                return {"status": "FAIL", "reason": f"mismatch with _resolved_correct: {'; '.join(diffs[:5])}"}

        # No expected output — check resolved is valid xlsx (openable)
        import openpyxl
        try:
            wb = openpyxl.load_workbook(resolved, data_only=True)
            wb.close()
        except Exception as e:
            return {"status": "FAIL", "reason": f"output not valid xlsx: {e}"}

        return {"status": "PASS", "reason": "resolved OK (no _resolved_correct to compare)"}

    except Exception as e:
        return {"status": "FAIL", "reason": str(e)}
    finally:
        shutil.rmtree(work, ignore_errors=True)


def main():
    print("=" * 60)
    print("DataTable Merge Test Runner")
    print(f"Test dir: {TEST_DIR}")
    print("=" * 60)

    cases = find_test_cases()
    if not cases:
        print("\nNo test cases found!")
        return

    print(f"\nFound {len(cases)} test case(s)\n")

    results = OrderedDict()
    for prefix, roles in cases.items():
        print(f"  {prefix} ... ", end="", flush=True)
        r = run_one_test(prefix, roles)
        results[prefix] = r
        status_icon = "PASS" if r["status"] == "PASS" else "FAIL" if r["status"] == "FAIL" else "SKIP"
        print(f"{status_icon} ({r.get('reason', '')})")

    # Summary
    passed = sum(1 for r in results.values() if r["status"] == "PASS")
    failed = sum(1 for r in results.values() if r["status"] == "FAIL")
    skipped = sum(1 for r in results.values() if r["status"] == "SKIP")

    print(f"\n{'='*60}")
    print(f"Results: {passed} passed, {failed} failed, {skipped} skipped")
    if failed:
        print("FAILED:")
        for prefix, r in results.items():
            if r["status"] == "FAIL":
                print(f"  {prefix}: {r['reason']}")
    print("=" * 60)

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
