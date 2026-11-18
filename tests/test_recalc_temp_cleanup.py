# -*- coding: utf-8 -*-
"""Tests for recalc temp-copy cleanup: retry deletion and stale-file sweep."""
import os
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import excel_recalc
from excel_recalc import cleanup_stale_recalc_copies, remove_with_retry, schedule_deletion


# ---------------------------------------------------------------------------
# remove_with_retry
# ---------------------------------------------------------------------------

def test_remove_with_retry_deletes_existing_file(tmp_path):
    f = tmp_path / ".X_old_recalc_abc123.xlsx"
    f.write_bytes(b"data")
    assert remove_with_retry(str(f)) is True
    assert not f.exists()


def test_remove_with_retry_missing_file_is_success(tmp_path):
    assert remove_with_retry(str(tmp_path / "nope.xlsx")) is True


def test_remove_with_retry_retries_until_lock_released(tmp_path, monkeypatch):
    f = tmp_path / "locked.xlsx"
    f.write_bytes(b"data")

    real_remove = os.remove
    state = {"calls": 0}

    def flaky_remove(p):
        state["calls"] += 1
        if state["calls"] < 3:
            raise PermissionError("locked")
        return real_remove(p)

    monkeypatch.setattr(os, "remove", flaky_remove)
    monkeypatch.setattr(time, "sleep", lambda *_: None)

    assert remove_with_retry(str(f), attempts=5, delay=0) is True
    assert state["calls"] == 3
    assert not f.exists()


def test_remove_with_retry_gives_up_after_attempts(tmp_path, monkeypatch):
    f = tmp_path / "locked.xlsx"
    f.write_bytes(b"data")

    def always_locked(p):
        raise PermissionError("locked")

    monkeypatch.setattr(os, "remove", always_locked)
    monkeypatch.setattr(time, "sleep", lambda *_: None)

    assert remove_with_retry(str(f), attempts=3, delay=0) is False


# ---------------------------------------------------------------------------
# cleanup_stale_recalc_copies
# ---------------------------------------------------------------------------

def _make_file(directory, name, age_seconds=0):
    p = Path(directory) / name
    p.write_bytes(b"x")
    if age_seconds:
        # Only mtime can be backdated (os.utime); creation time cannot.
        old = time.time() - age_seconds
        os.utime(str(p), (old, old))
    return p


def _patch_ctime(monkeypatch, ages):
    """Fake creation times: {filename: age_seconds}, default fresh."""
    now = time.time()

    def fake_ctime(path):
        age = ages.get(Path(path).name, 0)
        return now - age

    monkeypatch.setattr(excel_recalc, "_file_ctime", fake_ctime)


def test_sweep_removes_old_recalc_copies(tmp_path, monkeypatch):
    stale = _make_file(tmp_path, ".Effect_old_recalc_a1b2c3.xlsx")
    stale_no_uid = _make_file(tmp_path, ".Effect_old_recalc.xlsx")
    stale_single = _make_file(tmp_path, ".Effect_recalc_d4e5f6a7.xlsx")
    _patch_ctime(monkeypatch, {p.name: 7200 for p in (stale, stale_no_uid, stale_single)})

    cleanup_stale_recalc_copies(str(tmp_path))

    assert not stale.exists()
    assert not stale_no_uid.exists()
    assert not stale_single.exists()


def test_sweep_uses_creation_time_not_mtime(tmp_path):
    # copy2() preserves the SOURCE mtime, so a just-created copy of a
    # long-unmodified workbook has an old mtime but a fresh ctime. It must
    # survive the sweep — it may belong to a concurrent run.
    fresh_copy = _make_file(tmp_path, ".Effect_old_recalc_a1b2c3.xlsx", age_seconds=7200)

    cleanup_stale_recalc_copies(str(tmp_path))

    assert fresh_copy.exists(), "age must be measured by creation time, not mtime"


def test_sweep_keeps_young_copies_and_unrelated_files(tmp_path, monkeypatch):
    young = _make_file(tmp_path, ".Effect_old_recalc_a1b2c3.xlsx")  # fresh
    user_file = _make_file(tmp_path, "Effect.xlsx")
    other_tmp = _make_file(tmp_path, ".Effect_backup_a1b2c3.xlsx")
    _patch_ctime(monkeypatch, {})

    cleanup_stale_recalc_copies(str(tmp_path))

    assert young.exists(), "young copies may belong to a concurrent run"
    assert user_file.exists()
    assert other_tmp.exists()


def test_sweep_ignores_locked_files(tmp_path, monkeypatch):
    stale = _make_file(tmp_path, ".Effect_new_recalc_a1b2c3.xlsx")

    real_remove = os.remove

    def locked_remove(p):
        if "new_recalc" in str(p):
            raise PermissionError("locked")
        return real_remove(p)

    monkeypatch.setattr(os, "remove", locked_remove)
    cleanup_stale_recalc_copies(str(tmp_path), min_age_seconds=0)
    assert stale.exists()


# ---------------------------------------------------------------------------
# schedule_deletion (background reaper for locked temp copies)
# ---------------------------------------------------------------------------

@pytest.mark.skipif(os.name != "nt", reason="background reaper is Windows-only")
def test_schedule_deletion_removes_file(tmp_path):
    f = tmp_path / ".X_old_recalc_zz99.xlsx"
    f.write_bytes(b"data")

    schedule_deletion(str(f), attempts=15, delay=1)

    deadline = time.time() + 20
    while f.exists() and time.time() < deadline:
        time.sleep(0.5)
    assert not f.exists()


def test_schedule_deletion_missing_file_is_noop(tmp_path):
    # Must not raise even when the file is already gone.
    schedule_deletion(str(tmp_path / "nope.xlsx"), attempts=1, delay=1)


# ---------------------------------------------------------------------------
# recalc_to_temp sweeps before creating a new copy
# ---------------------------------------------------------------------------

def test_recalc_to_temp_sweeps_stale_copies(tmp_path, monkeypatch):
    src = _make_file(tmp_path, "Effect.xlsx")
    stale = _make_file(tmp_path, ".Effect_recalc_ffffffff.xlsx")
    _patch_ctime(monkeypatch, {stale.name: 7200})

    monkeypatch.setattr(excel_recalc, "recalc_xlsx", lambda *a, **k: a[0])
    copy_path = Path(excel_recalc.recalc_to_temp(str(src), str(tmp_path)))
    try:
        assert copy_path.exists()
        assert copy_path.name.startswith(".Effect_recalc_")
        assert not stale.exists(), "stale copies must be swept"
    finally:
        remove_with_retry(str(copy_path))


# ---------------------------------------------------------------------------
# diff_two_files leaves no temp copies behind
# ---------------------------------------------------------------------------

def _make_table(path, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "#TestSheet"
    for i, h in enumerate(["RowName", "Name", "Value"], 1):
        ws.cell(row=1, column=i, value=h)
        ws.cell(row=2, column=i, value=h)
        ws.cell(row=3, column=i, value="FString")
    for ri, row_data in enumerate(rows, 4):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val)
    wb.save(str(path))
    wb.close()


def test_diff_two_files_cleans_up_temp_copies(tmp_path, monkeypatch):
    import diff_engine

    old_f = tmp_path / "Effect.xlsx"
    new_f = tmp_path / "Effect_new.xlsx"
    _make_table(old_f, [("A", "Alice", "100")])
    _make_table(new_f, [("A", "Alice", "200"), ("B", "Bob", "300")])

    # Stale leftover from a "previous crashed run" — must be swept.
    stale = _make_file(tmp_path, ".Effect_old_recalc_deadbe.xlsx")
    # Fresh leftover simulating a concurrent run — must survive.
    young = _make_file(tmp_path, ".Effect_new_recalc_alive1.xlsx")
    _patch_ctime(monkeypatch, {stale.name: 7200})

    out_dir = tmp_path / ".temp"

    monkeypatch.setattr(excel_recalc, "recalc_batch", lambda paths, **k: paths)
    monkeypatch.setattr(os, "startfile", lambda *_: None)

    diff_engine.diff_two_files(str(old_f), str(new_f), str(out_dir),
                               display_path=str(new_f))

    leftovers = [p.name for p in tmp_path.glob(".*_recalc*.xlsx")]
    assert leftovers == [young.name], f"temp copies leaked: {leftovers}"
    assert not stale.exists()
    assert list(out_dir.glob("Effect*_diff_*.xlsx")), "diff output missing"


def test_diff_two_files_schedules_deletion_when_copies_locked(tmp_path, monkeypatch):
    """If a temp copy is still locked after the retry window, the background
    reaper must be handed the file instead of leaking it into the directory."""
    import diff_engine

    old_f = tmp_path / "Effect.xlsx"
    new_f = tmp_path / "Effect_new.xlsx"
    _make_table(old_f, [("A", "Alice", "100")])
    _make_table(new_f, [("A", "Alice", "200")])

    out_dir = tmp_path / ".temp"

    monkeypatch.setattr(excel_recalc, "recalc_batch", lambda paths, **k: paths)
    monkeypatch.setattr(os, "startfile", lambda *_: None)
    # diff_two_files imports remove_with_retry locally from excel_recalc.
    monkeypatch.setattr(excel_recalc, "remove_with_retry", lambda *a, **k: False)
    scheduled = []
    monkeypatch.setattr(diff_engine, "schedule_deletion", scheduled.append)

    diff_engine.diff_two_files(str(old_f), str(new_f), str(out_dir),
                               display_path=str(new_f))

    assert len(scheduled) == 2, f"reaper not engaged: {scheduled}"
    assert all("_recalc_" in p for p in scheduled)
