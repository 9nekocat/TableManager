# -*- coding: utf-8 -*-
"""
Standalone diff tool for TortoiseSVN integration.

Usage:
    diff.exe "%base" "%mine"

TortoiseSVN passes two arguments:
    %base — old version temp file (from SVN)
    %mine — working copy file

The tool compares the two files, generates a diff Excel in
RAW_DIR/.temp/, and opens it silently.
"""

import sys
import os
from pathlib import Path

# 源码运行时把 src/ 自身加入模块搜索路径（打包后由 PyInstaller 处理）
sys.path.insert(0, str(Path(__file__).resolve().parent))


def _label_to_path(label: str) -> Path | None:
    """Extract a working-copy path from an SVN -L label if present."""
    if not label:
        return None
    path_text = label.split("\t", 1)[0].strip().strip('"')
    if not path_text:
        return None
    return Path(path_text)


def _parse_diff_args(args):
    """Support both TortoiseSVN and svn.exe external diff invocations.

    TortoiseSVN usually calls: diff.exe "%base" "%mine"
    svn.exe calls: diff.exe -u -L <old-label> -L <new-label> <old> <new>
    """
    labels = []
    paths = []
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "-L" and i + 1 < len(args):
            labels.append(args[i + 1])
            i += 2
            continue
        if arg.startswith("-"):
            i += 1
            continue
        paths.append(arg)
        i += 1

    existing_files = [p for p in paths if Path(p).is_file()]
    if len(existing_files) >= 2:
        old_path = existing_files[-2]
        new_path = existing_files[-1]
    elif len(paths) >= 2:
        old_path = paths[-2]
        new_path = paths[-1]
    else:
        raise ValueError("Usage: diff.exe <old_file> <new_file>")

    display_path = None
    if labels:
        display_path = _label_to_path(labels[-1])
    if display_path is None:
        display_path = Path(new_path)

    return str(old_path), str(new_path), display_path


def _find_raw_dir(*candidate_paths: Path) -> Path:
    """Derive DataTable/Raw directory from the working copy file path.
    Priority:
      1. Walk up from mine_path to find DataTable/Raw
      2. CATSINTHESHELL_ROOT env var
      3. A:/DataTable/Raw (project convention)
    """
    # 1. Walk up from likely working-copy paths.
    for candidate in candidate_paths:
        if candidate is None:
            continue
        current = candidate if candidate.is_dir() else candidate.parent
        for _ in range(10):
            if current.name == "Raw" and current.parent.name == "DataTable":
                return current
            if current.parent == current:
                break
            current = current.parent

    # 2. Environment variable
    env_root = os.environ.get("CATSINTHESHELL_ROOT")
    if env_root:
        return Path(env_root) / "DataTable" / "Raw"

    # 3. Project convention
    fallback = Path("A:/DataTable/Raw")
    if fallback.exists():
        return fallback

    raise FileNotFoundError(
        "無法找到 DataTable/Raw 目錄。"
        "請設置環境變量 CATSINTHESHELL_ROOT 指向項目根目錄。"
    )


def main():
    if len(sys.argv) < 3:
        print(f"Usage: {sys.argv[0]} <old_file> <new_file>")
        sys.exit(1)

    try:
        old_path, new_path, display_path = _parse_diff_args(sys.argv[1:])
        raw_dir = _find_raw_dir(display_path, Path(new_path), Path(old_path))
        temp_dir = raw_dir / ".temp"

        # Set up env so _find_project_root in manager_core can find the project
        # (force-overwrite: a stale inherited value would write diff output
        # into the wrong project's .temp directory)
        os.environ["CATSINTHESHELL_ROOT"] = str(raw_dir.parent.parent)

        from diff_engine import diff_two_files

        diff_two_files(old_path, new_path, str(temp_dir), display_path=str(display_path))
    except Exception as e:
        # TortoiseSVN captures stderr — report error briefly
        print(f"[Diff Error] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
