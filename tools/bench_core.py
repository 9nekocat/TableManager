"""Synthetic benchmark for the pure diff core.

Usage:
    python tools/bench_core.py --rows 20000 --columns 20
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from core.diff import diff_documents
from core.model import Cell, Header, Row, Sheet, WorkbookDocument


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=20_000)
    parser.add_argument("--columns", type=int, default=20)
    parser.add_argument("--changes", type=int, default=100)
    args = parser.parse_args()

    headers = [Header(f"列{i}", f"Field{i}", "string") for i in range(args.columns)]
    rows = []
    for row_index in range(args.rows):
        rows.append(
            Row(
                f"ROW_{row_index:07d}",
                [Cell(value=f"{row_index}:{column}") for column in range(args.columns)],
            )
        )
    old_doc = WorkbookDocument({"#Data": Sheet("#Data", headers, rows)})
    new_rows = [
        Row(row.key, [cell.clone() for cell in row.cells], row.source_row)
        for row in rows
    ]
    for index in range(min(args.changes, len(new_rows))):
        new_rows[index].cells[1].value = f"changed-{index}"
    new_doc = WorkbookDocument({"#Data": Sheet("#Data", headers, new_rows)})

    started = time.perf_counter()
    result = diff_documents(old_doc, new_doc)
    elapsed = time.perf_counter() - started
    print(
        f"rows={args.rows} columns={args.columns} "
        f"changed_cells={result.summary['cells_modified']} "
        f"seconds={elapsed:.4f} cells/s={args.rows * args.columns / elapsed:,.0f}"
    )
    print(result.summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
