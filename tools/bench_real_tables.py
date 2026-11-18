"""Read-only benchmark for the refactored core using real project tables.

Raw workbooks are never saved. Temporary copies and generated benchmark
artifacts live outside ``DataTable/Raw``.
"""

from __future__ import annotations

import argparse
import cProfile
import gc
import json
import pstats
import shutil
import statistics
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Callable

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from core.diff import diff_documents, diff_files, format_text_diff, write_diff_excel
from core.merge import format_merge_text, merge_documents, merge_files, write_merge_excel
from core.model import WorkbookDocument
from core.text import workbook_to_text


DEFAULT_FILES = (
    "Localization/Card.xlsx",
    "Effect.xlsx",
    "Localization/Buff.xlsx",
    "ProjectileTable.xlsx",
)


def _measure(callable_: Callable[[], Any], runs: int) -> dict[str, Any]:
    times: list[float] = []
    value = None
    for _ in range(runs):
        gc.collect()
        gc.disable()
        started = time.perf_counter()
        try:
            value = callable_()
        finally:
            elapsed = time.perf_counter() - started
            gc.enable()
        times.append(elapsed)
    return {
        "median_s": statistics.median(times),
        "min_s": min(times),
        "max_s": max(times),
        "runs": times,
    }


def _mutate_cells(document: WorkbookDocument, change_count: int) -> int:
    cells = [
        cell
        for sheet in document.sheets.values()
        for row in sheet.rows
        for cell in row.cells[1:]
    ]
    if not cells:
        return 0
    target = min(change_count, len(cells))
    step = max(1, len(cells) // target)
    changed = 0
    for index in range(0, len(cells), step):
        if changed >= target:
            break
        cells[index].value = f"benchmark-change-{changed}"
        cells[index].formula = None
        cells[index].is_array = False
        changed += 1
    return changed


def _document_stats(document: WorkbookDocument) -> dict[str, int]:
    rows = sum(len(sheet.rows) for sheet in document.sheets.values())
    columns = sum(len(sheet.headers) for sheet in document.sheets.values())
    cells = sum(
        len(row.cells)
        for sheet in document.sheets.values()
        for row in sheet.rows
    )
    formulas = sum(
        bool(cell.formula)
        for sheet in document.sheets.values()
        for row in sheet.rows
        for cell in row.cells
    )
    return {
        "data_sheets": len(document.sheets),
        "rows": rows,
        "columns": columns,
        "cells": cells,
        "formulas": formulas,
        "load_issues": len(document.load_issues),
    }


def _benchmark_one(
    source: Path,
    local_copy: Path,
    output_dir: Path,
    *,
    runs: int,
    changes: int,
    profile: bool,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "file": source.name,
        "relative_path": source.as_posix(),
        "size_bytes": source.stat().st_size,
        "local_copy": str(local_copy),
    }
    result["a_read_bytes"] = _measure(lambda: source.read_bytes(), runs)
    result["local_read_bytes"] = _measure(lambda: local_copy.read_bytes(), runs)

    cold_started = time.perf_counter()
    cold_document = WorkbookDocument.from_xlsx(source)
    result["cold_load_a_s"] = time.perf_counter() - cold_started
    result["stats"] = _document_stats(cold_document)

    result["load_a_core"] = _measure(
        lambda: WorkbookDocument.from_xlsx(source),
        runs,
    )
    result["load_local_core"] = _measure(
        lambda: WorkbookDocument.from_xlsx(local_copy),
        runs,
    )

    old_document = WorkbookDocument.from_xlsx(local_copy)
    new_document = old_document.clone()
    changed = _mutate_cells(new_document, changes)
    result["changed_cells"] = changed
    result["validate_old"] = _measure(old_document.validate, runs)
    result["validate_new"] = _measure(new_document.validate, runs)
    result["text_snapshot_old"] = _measure(
        lambda: workbook_to_text(old_document),
        runs,
    )
    result["text_snapshot_new"] = _measure(
        lambda: workbook_to_text(new_document),
        runs,
    )
    diff_result = diff_documents(old_document, new_document)
    result["diff_model_with_validation"] = _measure(
        lambda: diff_documents(old_document, new_document),
        runs,
    )
    result["diff_model"] = _measure(
        lambda: diff_documents(
            old_document,
            new_document,
            validate_inputs=False,
        ),
        runs,
    )
    result["format_text"] = _measure(
        lambda: format_text_diff(diff_result),
        runs,
    )
    excel_path = output_dir / f"{source.stem}_diff_bench.xlsx"
    result["write_diff_excel"] = _measure(
        lambda: write_diff_excel(diff_result, excel_path),
        runs,
    )

    base = old_document.clone()
    ours = old_document.clone()
    theirs = old_document.clone()
    _mutate_cells(ours, max(1, changes // 2))
    other_cells = [
        cell
        for sheet in theirs.sheets.values()
        for row in sheet.rows
        for cell in row.cells[2:]
    ]
    for index, cell in enumerate(other_cells[: max(1, changes // 2)]):
        cell.value = f"benchmark-theirs-{index}"
        cell.formula = None
        cell.is_array = False
    merge_result = merge_documents(base, ours, theirs, strategy="ours")
    result["merge_with_validation"] = _measure(
        lambda: merge_documents(base, ours, theirs, strategy="ours"),
        runs,
    )
    result["merge_model"] = _measure(
        lambda: merge_documents(
            base,
            ours,
            theirs,
            strategy="ours",
            validate_output=False,
        ),
        runs,
    )
    result["format_merge_text"] = _measure(
        lambda: format_merge_text(merge_result),
        runs,
    )
    merge_path = output_dir / f"{source.stem}_merge_bench.xlsx"
    result["write_merge_excel"] = _measure(
        lambda: write_merge_excel(
            merge_result,
            merge_path,
            allow_conflicts=bool(merge_result.conflicts),
        ),
        runs,
    )
    result["diff_end_to_end_a"] = _measure(
        lambda: diff_files(source, source),
        runs,
    )
    result["merge_end_to_end_a"] = _measure(
        lambda: merge_files(source, source, source, strategy="ours"),
        runs,
    )

    if profile:
        profiler = cProfile.Profile()
        profiler.enable()
        profile_document = WorkbookDocument.from_xlsx(local_copy)
        profile_new = profile_document.clone()
        _mutate_cells(profile_new, changes)
        diff_documents(profile_document, profile_new)
        profiler.disable()
        stats = pstats.Stats(profiler)
        result["profile_top"] = [
            {
                "function": function,
                "cumulative_s": data[3],
                "self_s": data[2],
                "calls": data[1],
            }
            for function, data in sorted(
                stats.stats.items(),
                key=lambda item: item[1][3],
                reverse=True,
            )[:20]
        ]
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw-root", type=Path, default=Path("A:/DataTable/Raw"))
    parser.add_argument("--files", nargs="*", default=list(DEFAULT_FILES))
    parser.add_argument("--runs", type=int, default=3)
    parser.add_argument("--changes", type=int, default=100)
    parser.add_argument("--profile", action="store_true")
    parser.add_argument("--json-output", type=Path)
    args = parser.parse_args()

    if not args.raw_root.exists():
        raise SystemExit(f"Raw root not found: {args.raw_root}")
    temp_root = Path(tempfile.mkdtemp(prefix="tablemanager_bench_"))
    copy_root = temp_root / "local_copies"
    output_root = temp_root / "outputs"
    copy_root.mkdir()
    output_root.mkdir()
    results = {
        "raw_root": str(args.raw_root),
        "temp_root": str(temp_root),
        "runs": args.runs,
        "changes": args.changes,
        "files": [],
    }
    try:
        for relative_name in args.files:
            source = args.raw_root / relative_name
            if not source.exists():
                results["files"].append(
                    {"relative_path": relative_name, "error": "not found"}
                )
                continue
            local_copy = copy_root / relative_name.replace("/", "__")
            shutil.copy2(source, local_copy)
            print(f"[bench] {relative_name}", flush=True)
            results["files"].append(
                _benchmark_one(
                    source,
                    local_copy,
                    output_root,
                    runs=args.runs,
                    changes=args.changes,
                    profile=args.profile,
                )
            )
    finally:
        print(f"[bench] artifacts: {temp_root}", flush=True)

    text = json.dumps(results, ensure_ascii=False, indent=2, default=str)
    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(text, encoding="utf-8")
        print(f"[bench] json: {args.json_output}", flush=True)
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
