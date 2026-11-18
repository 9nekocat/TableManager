# -*- coding: utf-8 -*-
"""Transactional merge write-back helpers.

This module keeps the destructive part of a merge in one place: create a
byte-for-byte backup, apply the preview, refresh Excel formula caches, and
restore the backup if any step fails.
"""

from __future__ import annotations

import os
import shutil
import uuid
from datetime import datetime
from pathlib import Path

import openpyxl

from manager_core import RAW_DIR, TEMP_DIR, cleanup_conflict_artifacts
from merge_engine import (
    resolve_merge,
    _read_external_link_cache,
    normalize_external_link_targets_to_preferred_drive,
)
from excel_recalc import com_available, count_formula_error_caches, verify_formula_cache_guard
from svn_client import SvnClient


BACKUP_DIR = TEMP_DIR / "merge_backups"


class MergeWorkflowError(RuntimeError):
    """Raised when a transactional merge fails after backup handling."""

    def __init__(
        self,
        message,
        *,
        backup_path=None,
        restored=False,
        restore_error=None,
        original_error=None,
    ):
        super().__init__(message)
        self.backup_path = str(backup_path) if backup_path else None
        self.restored = restored
        self.restore_error = str(restore_error) if restore_error else None
        self.original_error = str(original_error) if original_error else message

    def to_dict(self):
        return {
            "error": str(self),
            "backup_path": self.backup_path,
            "restored": self.restored,
            "restore_error": self.restore_error,
            "original_error": self.original_error,
        }


def _as_posix(path):
    return str(path).replace("\\", "/")


def _relative_table_path(path):
    path = Path(path)
    try:
        return path.resolve().relative_to(RAW_DIR.resolve())
    except ValueError:
        return Path(path.name)


def _atomic_copy(src, dst):
    src = Path(src)
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(f".{dst.name}.{uuid.uuid4().hex}.tmp")
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def create_merge_backup(target_path, backup_root=None):
    """Copy the current target workbook into .temp/merge_backups."""
    target = Path(target_path)
    if not target.exists():
        return None

    root = Path(backup_root) if backup_root else BACKUP_DIR
    rel = _relative_table_path(target)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{rel.stem}.{stamp}.{uuid.uuid4().hex[:8]}.bak{rel.suffix}"
    backup_path = root / rel.parent / backup_name
    _atomic_copy(target, backup_path)
    return backup_path


def restore_merge_backup(backup_path, target_path):
    """Restore a workbook backup over the target path."""
    if not backup_path:
        return False
    _atomic_copy(Path(backup_path), Path(target_path))
    return True


def list_merge_backups_for(target_path, backup_root=None):
    target = Path(target_path)
    root = Path(backup_root) if backup_root else BACKUP_DIR
    rel = _relative_table_path(target)
    folder = root / rel.parent
    if not folder.exists():
        return []
    pattern = f"{rel.stem}.*.bak{rel.suffix}"
    return sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)


def _is_valid_workbook(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        wb.close()
        return True
    except Exception:
        return False


def latest_merge_backup_for(target_path, backup_root=None, require_valid=False):
    for backup in list_merge_backups_for(target_path, backup_root=backup_root):
        if not require_valid or _is_valid_workbook(backup):
            return backup
    return None


def find_merge_preview(query, temp_dir=None):
    """Find a *_merge.xlsx preview by exact path or fuzzy table name."""
    root = Path(temp_dir) if temp_dir else TEMP_DIR
    q = _as_posix(query).strip()
    candidates = []

    if q.lower().endswith("_merge.xlsx"):
        candidates.append(root / q)
    elif q.lower().endswith(".xlsx"):
        candidates.append(root / f"{q[:-5]}_merge.xlsx")
    else:
        candidates.append(root / f"{q}_merge.xlsx")

    for candidate in candidates:
        if candidate.exists():
            return candidate

    if not root.exists():
        return None

    needle = q.lower().replace("\\", "/")
    matches = []
    for path in root.rglob("*_merge.xlsx"):
        rel = _as_posix(path.relative_to(root))
        if needle in rel.lower() or needle in path.stem.lower():
            matches.append(path)
    return sorted(matches, key=lambda p: _as_posix(p.relative_to(root)).lower())[0] if matches else None


def original_name_from_preview(preview_path, temp_dir=None):
    root = Path(temp_dir) if temp_dir else TEMP_DIR
    preview = Path(preview_path)
    try:
        rel = preview.relative_to(root)
    except ValueError:
        rel = Path(preview.name)
    rel_text = _as_posix(rel)
    if rel_text.lower().endswith("_merge.xlsx"):
        return f"{rel_text[:-11]}.xlsx"
    return rel_text


def infer_theirs_path(original_name):
    svn = SvnClient()
    conflicts = svn.find_conflict_files()
    key = _as_posix(original_name)
    if key in conflicts:
        return conflicts[key].get("theirs")
    return None



def _external_link_files(path):
    cache = _read_external_link_cache(path)
    if not cache:
        return {}
    return cache.get("files", {}) or {}


def _verify_external_links_preserved(source_path, result_path):
    before = _external_link_files(source_path)
    if not before:
        return
    after = _external_link_files(result_path)
    missing = sorted(name for name in before if name not in after)
    if missing:
        raise RuntimeError(f"missing external links: {', '.join(missing[:5])}")


def _verify_external_links_unchanged(source_cache, result_path):
    before = (source_cache or {}).get("files", {}) or {}
    if not before:
        return
    after = _external_link_files(result_path)
    missing = sorted(name for name in before if name not in after)
    changed = sorted(name for name, content in before.items() if after.get(name) != content)
    if missing or changed:
        details = []
        if missing:
            details.append(f"missing external links: {', '.join(missing[:5])}")
        if changed:
            details.append(f"changed external links: {', '.join(changed[:5])}")
        raise RuntimeError("; ".join(details))


def apply_merge_preview_transactional(
    preview_path,
    output_path,
    *,
    mine_path=None,
    theirs_path=None,
    backup_root=None,
    mark_svn_resolved=False,
    cleanup=False,
):
    """Apply a merge preview and restore the original workbook on failure."""
    preview = Path(preview_path)
    output = Path(output_path)
    mine = Path(mine_path or output)
    backup_path = None

    try:
        backup_path = create_merge_backup(output, backup_root=backup_root)
        merge_base = mine
        try:
            if backup_path and mine.resolve() == output.resolve():
                merge_base = backup_path
        except OSError:
            pass
        # 写回前基线：备份完成后、resolve 前采集（基准 = 未修改的目标表）。
        # recover 场景目标表可能已损坏/缺失，无基线可采 → 空基线
        try:
            baseline = count_formula_error_caches(str(output))
        except Exception:
            baseline = {}
        result = resolve_merge(
            str(preview),
            str(output),
            mine_path=str(merge_base),
            theirs_path=theirs_path,
        )
        if com_available():
            na_guard = verify_formula_cache_guard(result, baseline)
        else:
            na_guard = {
                "ok": None,
                "skipped": "本机无 win32com/Excel，已降级 openpyxl 写回，跳过 #N/A 基线校验",
            }
        link_targets = normalize_external_link_targets_to_preferred_drive(result)
        _verify_external_links_preserved(merge_base, result)
        svn_resolved = False
        if mark_svn_resolved:
            rel_path = _as_posix(output.relative_to(RAW_DIR))
            SvnClient().resolved(rel_path)
            svn_resolved = True

        cleanup_result = {"deleted": [], "failed": []}
        if cleanup:
            cleanup_result = cleanup_conflict_artifacts(output)

        return {
            "ok": True,
            "output": str(result),
            "backup_path": str(backup_path) if backup_path else None,
            "na_guard": na_guard,
            "external_link_targets": link_targets,
            "svn_resolved": svn_resolved,
            "cleanup": cleanup_result,
            "restored": False,
        }
    except Exception as exc:
        restored = False
        restore_error = None
        if backup_path:
            try:
                restored = restore_merge_backup(backup_path, output)
            except Exception as err:
                restore_error = err
        message = str(exc)
        if restored:
            message = f"{message}; original workbook restored from backup"
        elif backup_path:
            message = f"{message}; failed to restore backup"
        raise MergeWorkflowError(
            message,
            backup_path=backup_path,
            restored=restored,
            restore_error=restore_error,
            original_error=exc,
        ) from exc


def recover_merge_preview_transactional(
    preview_path,
    output_path,
    *,
    backup_path=None,
    theirs_path=None,
    backup_root=None,
):
    """Recover table changes from an existing *_merge.xlsx preview.

    If a valid backup exists, it is used as the clean mine/base workbook so this
    can repair a target workbook that was damaged by a previous failed write.
    """
    output = Path(output_path)
    clean_base = Path(backup_path) if backup_path else latest_merge_backup_for(
        output, backup_root=backup_root, require_valid=True
    )
    if clean_base is None:
        clean_base = output

    result = apply_merge_preview_transactional(
        preview_path,
        output,
        mine_path=str(clean_base),
        theirs_path=theirs_path,
        backup_root=backup_root,
        mark_svn_resolved=False,
        cleanup=False,
    )
    result["recovery_base"] = str(clean_base)
    return result
