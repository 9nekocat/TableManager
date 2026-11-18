"""Benchmark one-pass load, validation, and self-diff for every Raw xlsx."""

from __future__ import annotations

import argparse
import gc
import json
import statistics
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from core.diff import diff_documents
from core.model import WorkbookDocument


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw-root", type=Path, default=Path("A:/DataTable/Raw"))
    parser.add_argument("--runs", type=int, default=2)
    parser.add_argument("--json-output", type=Path)
    args = parser.parse_args()

    files = sorted(
        path for path in args.raw_root.rglob("*.xlsx")
        if ".temp" not in {part.casefold() for part in path.parts}
    )
    runs: list[dict[str, object]] = []
    for run_index in range(args.runs):
        gc.collect()
        per_file = []
        total_started = time.perf_counter()
        for path in files:
            started = time.perf_counter()
            try:
                document = WorkbookDocument.from_xlsx(path)
            except Exception as exc:
                load_s = time.perf_counter() - started
                per_file.append(
                    {
                        "file": str(path.relative_to(args.raw_root)),
                        "load_s": load_s,
                        "validate_s": 0.0,
                        "diff_s": 0.0,
                        "sheets": 0,
                        "issues": 0,
                        "error": f"{type(exc).__name__}: {exc}",
                    }
                )
                continue
            load_s = time.perf_counter() - started
            started = time.perf_counter()
            report = document.validate()
            validate_s = time.perf_counter() - started
            started = time.perf_counter()
            diff_documents(
                document,
                document.clone(),
                validate_inputs=False,
            )
            diff_s = time.perf_counter() - started
            per_file.append(
                {
                    "file": str(path.relative_to(args.raw_root)),
                    "load_s": load_s,
                    "validate_s": validate_s,
                    "diff_s": diff_s,
                    "sheets": len(document.sheets),
                    "issues": len(report.issues),
                    "error": None,
                }
            )
        elapsed = time.perf_counter() - total_started
        runs.append(
            {
                "run": run_index + 1,
                "total_s": elapsed,
                "files": per_file,
            }
        )
        print(f"[corpus] run={run_index + 1} seconds={elapsed:.3f}", flush=True)

    summary = {
        "raw_root": str(args.raw_root),
        "file_count": len(files),
        "runs": runs,
        "median_total_s": statistics.median(run["total_s"] for run in runs),
        "median_load_s": statistics.median(
            sum(item["load_s"] for item in run["files"]) for run in runs
        ),
        "median_validate_s": statistics.median(
            sum(item["validate_s"] for item in run["files"]) for run in runs
        ),
        "median_diff_s": statistics.median(
            sum(item["diff_s"] for item in run["files"]) for run in runs
        ),
    }
    text = json.dumps(summary, ensure_ascii=False, indent=2)
    if args.json_output:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(text, encoding="utf-8")
        print(f"[corpus] json: {args.json_output}", flush=True)
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
