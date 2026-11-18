# -*- coding: utf-8 -*-
"""
Reproducible test: external links survive merge resolve round-trip.

Usage:
    cd A:/Tools/TableManager
    python tests/test_external_links.py
"""

import hashlib
import os
import shutil
import tempfile
import zipfile
import sys
from pathlib import Path

# Ensure project root is on path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import merge_engine as merge_engine_module
from merge_engine import (
    generate_merge_preview,
    resolve_merge,
    _inject_external_links,
    normalize_external_link_targets_to_preferred_drive,
)
import openpyxl

# --- Config ---
# Use Localization/Card.xlsx — 4 external links (Effect.xlsx etc.), 8 zip entries
TEST_TABLE = "Localization/Card.xlsx"
# Number of external link files expected in xlsx zip
EXPECTED_EXTLINK_FILES = 8


def test_external_link_targets_prefer_a_drive(tmp_path, monkeypatch):
    monkeypatch.setattr(merge_engine_module, "RAW_DIR", tmp_path)
    xlsx_path = tmp_path / "Localization" / "Card.xlsx"
    xlsx_path.parent.mkdir(parents=True)
    rels_name = "xl/externalLinks/_rels/externalLink1.xml.rels"
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
        'Target="file:///D:\\work\\DataTable\\Raw\\Localization\\AttributeTable.xlsx" TargetMode="External"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
        'Target="/DataTable/Raw/Effect.xlsx" TargetMode="External"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
        'Target="Tag.xlsx" TargetMode="External"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(rels_name, rels_xml)

    result = normalize_external_link_targets_to_preferred_drive(xlsx_path, preferred_drive="A")

    assert result["rewritten"] == 3
    with zipfile.ZipFile(xlsx_path, "r") as z:
        rewritten = z.read(rels_name).decode("utf-8")
    assert 'Target="file:///A:\\DataTable\\Raw\\Localization\\AttributeTable.xlsx"' in rewritten
    assert 'Target="file:///A:\\DataTable\\Raw\\Effect.xlsx"' in rewritten
    assert 'Target="file:///A:\\DataTable\\Raw\\Localization\\Tag.xlsx"' in rewritten


def sha256_hashes(zip_path):
    """Return {filename: sha256hex} for all entries in a zip."""
    result = {}
    with zipfile.ZipFile(zip_path, 'r') as z:
        for n in z.namelist():
            result[n] = hashlib.sha256(z.read(n)).hexdigest()
    return result


def test_external_links_survive_round_trip():
    """Test that external link XML files survive load+save unchanged."""
    from manager_core import RAW_DIR

    src = RAW_DIR / TEST_TABLE
    if not src.exists():
        print(f"SKIP: {src} not found (need test data with external links)")
        return True  # not a failure

    orig_hashes = sha256_hashes(str(src))
    extlink_orig = [n for n in orig_hashes if n.startswith("xl/externalLinks/")]
    assert len(extlink_orig) >= 4, f"Expected >=4 extlink files, got {len(extlink_orig)}"

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        pass

    try:
        wb = openpyxl.load_workbook(str(src), data_only=False)
        wb.save(tmp.name)

        saved_hashes = sha256_hashes(tmp.name)
        extlink_saved = [n for n in saved_hashes if n.startswith("xl/externalLinks/")]

        # All extlink files should exist but WILL be corrupted
        corrupted = 0
        for n in extlink_orig:
            if n not in saved_hashes or orig_hashes[n] != saved_hashes.get(n):
                corrupted += 1

        print(f"  Round-trip without fix: {corrupted}/{len(extlink_orig)} extlink files corrupted")
        assert corrupted > 0, "Expected corruption — if none, openpyxl changed behavior"

        # Now apply the fix
        _inject_external_links(str(src), tmp.name)
        fixed_hashes = sha256_hashes(tmp.name)

        for n in extlink_orig:
            assert n in fixed_hashes, f"Missing after inject: {n}"
            assert fixed_hashes[n] == orig_hashes[n], f"Byte mismatch: {n}"

        print(f"  After _inject_external_links: 0/{len(extlink_orig)} corrupted — PASS")
    finally:
        os.unlink(tmp.name)

    return True


def test_merge_resolve_preserves_external_links():
    """Full merge pipeline: preview + resolve, verify external links intact."""
    from manager_core import RAW_DIR

    src = RAW_DIR / TEST_TABLE
    if not src.exists():
        print(f"SKIP: {src} not found")
        return True

    orig_hashes = sha256_hashes(str(src))
    extlink_orig = [n for n in orig_hashes if n.startswith("xl/externalLinks/")]
    assert len(extlink_orig) >= 4, f"Expected >=4 extlink files"

    work_dir = tempfile.mkdtemp()
    try:
        base = os.path.join(work_dir, "base.xlsx")
        mine = os.path.join(work_dir, "mine.xlsx")
        theirs = os.path.join(work_dir, "theirs.xlsx")
        preview = os.path.join(work_dir, "preview_merge.xlsx")
        resolved = os.path.join(work_dir, "resolved.xlsx")

        shutil.copy2(str(src), base)
        shutil.copy2(str(src), mine)
        shutil.copy2(str(src), theirs)

        # Modify one cell in "mine" to create a conflict
        wb = openpyxl.load_workbook(mine, data_only=False)
        data_sheets = [s for s in wb.sheetnames if s.startswith("#")]
        if data_sheets:
            ws = wb[data_sheets[0]]
            ws.cell(row=4, column=3, value="MINE_MOD_VALUE")
        wb.save(mine)
        # openpyxl save corrupts external links in mine — restore from pristine source
        _inject_external_links(str(src), mine)

        # Modify a different cell in "theirs" to create a real conflict
        wb = openpyxl.load_workbook(theirs, data_only=False)
        data_sheets = [s for s in wb.sheetnames if s.startswith("#")]
        if data_sheets:
            ws = wb[data_sheets[0]]
            ws.cell(row=5, column=3, value="THEIRS_MOD_VALUE")
        wb.save(theirs)

        # Step 1: Generate merge preview
        result = generate_merge_preview(base, mine, theirs, preview)
        print(f"  Preview generated: {result}")

        # Step 2: Resolve merge — this writes back to mine (resolved path)
        result = resolve_merge(preview, resolved, mine_path=mine)
        print(f"  Resolved: {result}")

        # Step 3: Verify resolved file
        resolved_hashes = sha256_hashes(resolved)
        extlink_resolved = [n for n in resolved_hashes if n.startswith("xl/externalLinks/")]

        mismatches = 0
        for n in extlink_orig:
            if n not in resolved_hashes:
                print(f"  FAIL: {n} missing after merge resolve")
                mismatches += 1
            elif resolved_hashes[n] != orig_hashes[n]:
                print(f"  FAIL: {n} byte mismatch after merge resolve")
                mismatches += 1

        if mismatches == 0:
            print(f"  Full merge pipeline: {len(extlink_orig)}/{len(extlink_orig)} extlink files intact — PASS")
        else:
            print(f"  FAIL: {mismatches} extlink file(s) corrupted in full merge pipeline")
            return False

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    return True


def test_no_duplicate_entries():
    """Verify no duplicate zip entries after injection."""
    from manager_core import RAW_DIR

    src = RAW_DIR / TEST_TABLE
    if not src.exists():
        print(f"SKIP: {src} not found")
        return True

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        pass

    try:
        wb = openpyxl.load_workbook(str(src), data_only=False)
        wb.save(tmp.name)
        _inject_external_links(str(src), tmp.name)

        with zipfile.ZipFile(tmp.name, 'r') as z:
            from collections import Counter
            names = z.namelist()
            dupes = [n for n, c in Counter(names).items() if c > 1]

        if dupes:
            print(f"  FAIL: Duplicate entries: {dupes}")
            return False
        else:
            print(f"  No duplicate entries: PASS")
    finally:
        os.unlink(tmp.name)

    return True


if __name__ == "__main__":
    print("=" * 60)
    print("External Link Preservation Tests")
    print("=" * 60)

    all_pass = True

    print("\n[1] Round-trip load+save+inject")
    all_pass &= test_external_links_survive_round_trip()

    print("\n[2] Full merge pipeline (preview + resolve)")
    all_pass &= test_merge_resolve_preserves_external_links()

    print("\n[3] No duplicate zip entries")
    all_pass &= test_no_duplicate_entries()

    print("\n" + "=" * 60)
    if all_pass:
        print("ALL TESTS PASSED")
    else:
        print("SOME TESTS FAILED")
    print("=" * 60)
    sys.exit(0 if all_pass else 1)
