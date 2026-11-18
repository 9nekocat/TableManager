"""Command line interface for the refactored table core."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from core.diff import diff_files, format_text_diff, write_diff_excel
from core.merge import (
    MergeResult,
    format_merge_text,
    merge_files,
    write_merge_excel,
)
from core.model import WorkbookDocument
from core.operations import apply_operations


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="table-core",
        description="DataTable core: validate, diff, merge, and CRUD.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    validate = subparsers.add_parser("validate", help="Validate one workbook")
    validate.add_argument("file")
    validate.add_argument("--json", action="store_true")
    validate.add_argument("--clean", action="store_true", help="Report and remove blank rows/columns")

    inspect_parser = subparsers.add_parser("inspect", help="Inspect workbook structure")
    inspect_parser.add_argument("file")
    inspect_parser.add_argument("--sheet")
    inspect_parser.add_argument("--json", action="store_true")

    diff = subparsers.add_parser("diff", help="Compare two workbooks")
    diff.add_argument("old")
    diff.add_argument("new")
    diff.add_argument("--format", choices=("text", "json", "excel"), default="text")
    diff.add_argument("--output")

    merge = subparsers.add_parser("merge", help="Three-way merge")
    merge.add_argument("base")
    merge.add_argument("ours")
    merge.add_argument("theirs")
    merge.add_argument("--output")
    merge.add_argument("--strategy", choices=("fail", "ours", "theirs"), default="fail")
    merge.add_argument("--report", choices=("text", "json"), default="text")
    merge.add_argument(
        "--allow-conflicts",
        action="store_true",
        help="Write the Excel preview even when unresolved conflicts exist",
    )

    edit = subparsers.add_parser("edit", help="Apply JSON CRUD operations")
    edit.add_argument("file")
    edit.add_argument("--ops", required=True, help="JSON file containing an operation list")
    edit.add_argument("--output", required=True)
    edit.add_argument("--json", action="store_true")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "validate":
        document = WorkbookDocument.from_xlsx(args.file)
        report = document.validate(clean_first=args.clean)
        if args.json:
            _print_json(report.to_dict())
        else:
            print(format_validation_text(report))
        return 0 if report.ok else 1

    if args.command == "inspect":
        document = WorkbookDocument.from_xlsx(args.file)
        sheets = (
            {args.sheet: document.sheet(args.sheet)}
            if args.sheet
            else document.sheets
        )
        data = {
            "file": str(args.file),
            "sheets": [
                {
                    "name": sheet.name,
                    "columns": [header.to_dict() for header in sheet.headers],
                    "rows": len(sheet.rows),
                    "row_keys": [row.key for row in sheet.rows[:20]],
                }
                for sheet in sheets.values()
            ],
            "issues": [issue.to_dict() for issue in document.load_issues],
        }
        if args.json:
            _print_json(data)
        else:
            for sheet in data["sheets"]:
                print(
                    f"{sheet['name']}: {sheet['rows']} rows, "
                    f"{len(sheet['columns'])} columns"
                )
                print("  columns: " + ", ".join(
                    column["en"] or column["cn"] for column in sheet["columns"]
                ))
                print("  row keys: " + ", ".join(sheet["row_keys"]))
            for issue in data["issues"]:
                print(f"[{issue['severity']}] {issue['code']}: {issue['message']}")
        return 0

    if args.command == "diff":
        result = diff_files(args.old, args.new)
        if args.format == "excel":
            output = args.output or str(Path(args.new).with_name(Path(args.new).stem + "_diff.xlsx"))
            print(write_diff_excel(result, output))
        elif args.format == "json":
            _print_json(result.to_dict())
        else:
            print(format_text_diff(result))
        return 0 if not result.issues or not any(issue.is_error for issue in result.issues) else 1

    if args.command == "merge":
        result = merge_files(
            args.base,
            args.ours,
            args.theirs,
            strategy=args.strategy,
        )
        if args.report == "json":
            _print_json(result.to_dict())
        else:
            print(format_merge_text(result))
        output = args.output or str(
            Path(args.ours).with_name(Path(args.ours).stem + "_merged.xlsx")
        )
        try:
            write_merge_excel(
                result,
                output,
                allow_conflicts=args.allow_conflicts,
            )
        except Exception as exc:
            if result.conflicts and not args.allow_conflicts:
                print(f"[merge] 未写入结果: {exc}", file=sys.stderr)
                return 2
            raise
        print(f"[merge] written: {output}")
        return 0 if result.ok else 2

    if args.command == "edit":
        document = WorkbookDocument.from_xlsx(args.file)
        with open(args.ops, "r", encoding="utf-8-sig") as handle:
            operations = json.load(handle)
        if not isinstance(operations, list):
            raise ValueError("操作文件必须是 JSON 数组")
        apply_operations(document, operations)
        report = document.validate(clean_first=True)
        document.save(args.output)
        if args.json:
            _print_json(report.to_dict())
        else:
            print(format_validation_text(report))
            print(f"[edit] written: {args.output}")
        return 0 if report.ok else 1

    return 1


def format_validation_text(report) -> str:
    summary = report.to_dict()["summary"]
    lines = [
        "validation: " + ("OK" if report.ok else "FAILED"),
        (
            f"sheets={summary['sheets']} rows={summary['rows']} "
            f"columns={summary['columns']} formulas={summary['formulas']} "
            f"errors={summary['errors']} warnings={summary['warnings']}"
        ),
    ]
    if summary["dropped_blank_rows"] or summary["dropped_blank_columns"]:
        lines.append(
            f"cleaned blank rows={summary['dropped_blank_rows']} "
            f"blank columns={summary['dropped_blank_columns']}"
        )
    for issue in report.issues:
        location = " / ".join(
            str(value)
            for value in (issue.sheet, issue.row, issue.column)
            if value is not None
        )
        lines.append(
            f"- [{issue.severity}] {issue.code}: {issue.message}"
            + (f" ({location})" if location else "")
        )
    return "\n".join(lines)


def _print_json(value: Any) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    raise SystemExit(main())
