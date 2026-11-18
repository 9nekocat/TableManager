"""Table diff over stable text snapshots.

Formula cells use Excel's cached values only. No formula AST comparison or
recalculation happens in this module.
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from .model import Header, Issue, WorkbookDocument
from .text import TextRow, TextSheet, workbook_to_text


@dataclass(frozen=True, slots=True)
class CellChange:
    column: str
    old: str
    new: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "column": self.column,
            "old": self.old,
            "new": self.new,
        }


@dataclass(slots=True)
class RowChange:
    sheet: str
    row: str
    kind: str
    old_row: TextRow | None = None
    new_row: TextRow | None = None
    cell_changes: list[CellChange] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "row": self.row,
            "kind": self.kind,
            "cells": [change.to_dict() for change in self.cell_changes],
        }


@dataclass(frozen=True, slots=True)
class ColumnChange:
    sheet: str
    column: str
    kind: str
    old: Header | None = None
    new: Header | None = None
    old_index: int | None = None
    new_index: int | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "column": self.column,
            "kind": self.kind,
            "old": self.old.to_dict() if self.old else None,
            "new": self.new.to_dict() if self.new else None,
            "old_index": self.old_index,
            "new_index": self.new_index,
        }


@dataclass(slots=True)
class SheetDiff:
    sheet: str
    status: str
    columns: tuple[Header, ...] = ()
    column_changes: list[ColumnChange] = field(default_factory=list)
    row_changes: list[RowChange] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return bool(self.column_changes or self.row_changes)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "status": self.status,
            "columns": [column.to_dict() for column in self.columns],
            "column_changes": [change.to_dict() for change in self.column_changes],
            "rows": [change.to_dict() for change in self.row_changes],
        }


@dataclass(slots=True)
class DiffResult:
    old_path: str | None
    new_path: str | None
    sheets: list[SheetDiff] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return any(sheet.has_changes for sheet in self.sheets)

    @property
    def summary(self) -> dict[str, int]:
        result = {
            "rows_added": 0,
            "rows_deleted": 0,
            "rows_modified": 0,
            "columns_added": 0,
            "columns_deleted": 0,
            "columns_modified": 0,
            "columns_moved": 0,
            "cells_modified": 0,
        }
        for sheet in self.sheets:
            for change in sheet.column_changes:
                key = {
                    "added": "columns_added",
                    "deleted": "columns_deleted",
                    "modified": "columns_modified",
                    "moved": "columns_moved",
                }.get(change.kind)
                if key:
                    result[key] += 1
            for change in sheet.row_changes:
                key = {
                    "added": "rows_added",
                    "deleted": "rows_deleted",
                    "modified": "rows_modified",
                }.get(change.kind)
                if key:
                    result[key] += 1
                result["cells_modified"] += len(change.cell_changes)
        return result

    def to_dict(self) -> dict[str, Any]:
        return {
            "old": self.old_path,
            "new": self.new_path,
            "comparison": "text_snapshot_with_cached_formula_values",
            "has_changes": self.has_changes,
            "summary": self.summary,
            "sheets": [sheet.to_dict() for sheet in self.sheets],
            "issues": [issue.to_dict() for issue in self.issues],
        }


@dataclass(frozen=True, slots=True)
class _AlignedColumn:
    key: str
    kind: str
    old_index: int | None
    new_index: int | None
    old_header: Header | None
    new_header: Header | None


def diff_files(
    old_path: str | Path,
    new_path: str | Path,
) -> DiffResult:
    old_doc = WorkbookDocument.from_xlsx(old_path)
    new_doc = WorkbookDocument.from_xlsx(new_path)
    result = diff_documents(old_doc, new_doc)
    result.old_path = str(old_path)
    result.new_path = str(new_path)
    return result


def diff_documents(
    old_doc: WorkbookDocument,
    new_doc: WorkbookDocument,
    *,
    validate_inputs: bool = True,
) -> DiffResult:
    result = DiffResult(
        old_path=old_doc.source_path,
        new_path=new_doc.source_path,
    )
    if validate_inputs:
        result.issues.extend(old_doc.validate().issues)
        result.issues.extend(new_doc.validate().issues)
    else:
        result.issues.extend(old_doc.load_issues)
        result.issues.extend(new_doc.load_issues)
    old_snapshot = workbook_to_text(old_doc)
    new_snapshot = workbook_to_text(new_doc)
    result.issues.extend(old_snapshot.issues)
    result.issues.extend(new_snapshot.issues)
    sheet_names = list(
        OrderedDict.fromkeys(
            [*old_snapshot.sheets.keys(), *new_snapshot.sheets.keys()]
        ).keys()
    )
    for sheet_name in sheet_names:
        old_sheet = old_snapshot.sheets.get(sheet_name)
        new_sheet = new_snapshot.sheets.get(sheet_name)
        if old_sheet is None:
            result.sheets.append(
                SheetDiff(
                    sheet_name,
                    "added",
                    columns=new_sheet.headers,
                    row_changes=_row_changes(None, new_sheet),
                )
            )
            continue
        if new_sheet is None:
            result.sheets.append(
                SheetDiff(
                    sheet_name,
                    "deleted",
                    columns=old_sheet.headers,
                    row_changes=_row_changes(old_sheet, None),
                )
            )
            continue
        result.sheets.append(_diff_sheet(old_sheet, new_sheet))
    return result


def _row_changes(old_sheet: TextSheet | None, new_sheet: TextSheet | None) -> list[RowChange]:
    if old_sheet is None and new_sheet is not None:
        return [
            RowChange(new_sheet.name, row.key, "added", new_row=row)
            for row in new_sheet.rows
        ]
    if new_sheet is None and old_sheet is not None:
        return [
            RowChange(old_sheet.name, row.key, "deleted", old_row=row)
            for row in old_sheet.rows
        ]
    return []


def _diff_sheet(
    old_sheet: TextSheet,
    new_sheet: TextSheet,
) -> SheetDiff:
    aligned, column_changes = _align_columns(old_sheet, new_sheet)
    row_changes = _diff_rows(
        old_sheet,
        new_sheet,
        aligned,
    )
    status = "modified" if column_changes or row_changes else "same"
    return SheetDiff(
        new_sheet.name,
        status,
        new_sheet.headers,
        column_changes,
        row_changes,
    )


def _align_columns(
    old_sheet: TextSheet,
    new_sheet: TextSheet,
) -> tuple[list[_AlignedColumn], list[ColumnChange]]:
    old_keys = list(old_sheet.column_keys)
    new_keys = list(new_sheet.column_keys)
    old_key_to_index = {key: index for index, key in enumerate(old_keys)}
    new_key_to_index = {key: index for index, key in enumerate(new_keys)}
    common = set(old_keys) & set(new_keys)
    paired_old: set[int] = set()
    paired_new: set[int] = set()
    aligned: list[_AlignedColumn] = []
    changes: list[ColumnChange] = []

    matcher = SequenceMatcher(a=old_keys, b=new_keys, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                old_index = i1 + offset
                new_index = j1 + offset
                key = old_keys[old_index]
                aligned.append(
                    _AlignedColumn(
                        key,
                        "same",
                        old_index,
                        new_index,
                        old_sheet.headers[old_index],
                        new_sheet.headers[new_index],
                    )
                )
                paired_old.add(old_index)
                paired_new.add(new_index)
        elif tag == "replace":
            old_count = i2 - i1
            new_count = j2 - j1
            pair_count = min(old_count, new_count)
            for offset in range(pair_count):
                old_index = i1 + offset
                new_index = j1 + offset
                old_header = old_sheet.headers[old_index]
                new_header = new_sheet.headers[new_index]
                key = old_keys[old_index]
                aligned.append(
                    _AlignedColumn(key, "modified", old_index, new_index, old_header, new_header)
                )
                changes.append(
                    ColumnChange(
                        old_sheet.name,
                        new_header.display_name(),
                        "modified",
                        old_header,
                        new_header,
                        old_index,
                        new_index,
                    )
                )
                paired_old.add(old_index)
                paired_new.add(new_index)
            for old_index in range(i1 + pair_count, i2):
                if old_index in paired_old:
                    continue
                header = old_sheet.headers[old_index]
                changes.append(
                    ColumnChange(
                        old_sheet.name,
                        header.display_name(),
                        "deleted",
                        old=header,
                        old_index=old_index,
                    )
                )
                paired_old.add(old_index)
            for new_index in range(j1 + pair_count, j2):
                if new_index in paired_new:
                    continue
                header = new_sheet.headers[new_index]
                changes.append(
                    ColumnChange(
                        old_sheet.name,
                        header.display_name(),
                        "added",
                        new=header,
                        new_index=new_index,
                    )
                )
                paired_new.add(new_index)

    for old_index, key in enumerate(old_keys):
        if old_index in paired_old or key not in common:
            continue
        new_index = new_key_to_index[key]
        aligned.append(
            _AlignedColumn(
                key,
                "same",
                old_index,
                new_index,
                old_sheet.headers[old_index],
                new_sheet.headers[new_index],
            )
        )
        paired_old.add(old_index)
        paired_new.add(new_index)

    common_old_order = [key for key in old_keys if key in common]
    common_new_order = [key for key in new_keys if key in common]
    moved = set()
    for tag, i1, i2, j1, j2 in SequenceMatcher(
        a=common_old_order, b=common_new_order, autojunk=False
    ).get_opcodes():
        if tag != "equal":
            moved.update(common_old_order[i1:i2])
            moved.update(common_new_order[j1:j2])
    for item in aligned:
        if (
            item.kind == "same"
            and item.old_index is not None
            and item.new_index is not None
            and old_keys[item.old_index] in moved
            and item.old_index != item.new_index
        ):
            changes.append(
                ColumnChange(
                    old_sheet.name,
                    item.new_header.display_name() if item.new_header else item.key,
                    "moved",
                    item.old_header,
                    item.new_header,
                    item.old_index,
                    item.new_index,
                )
            )

    # Keep the output order readable: current new order first, then deleted
    # columns in their old order.
    aligned.sort(
        key=lambda item: (
            0 if item.new_index is not None else 1,
            item.new_index if item.new_index is not None else item.old_index or 0,
        )
    )
    changes.sort(
        key=lambda item: (
            item.new_index if item.new_index is not None else 10_000 + (item.old_index or 0),
            item.kind,
        )
    )
    return aligned, changes


def _diff_rows(
    old_sheet: TextSheet,
    new_sheet: TextSheet,
    aligned: list[_AlignedColumn],
) -> list[RowChange]:
    old_rows = _row_identity_map(old_sheet)
    new_rows = _row_identity_map(new_sheet)
    identities = list(OrderedDict.fromkeys([*old_rows.keys(), *new_rows.keys()]))
    changes: list[RowChange] = []
    skipped_columns = {
        item.key for item in aligned if item.kind == "modified"
    }
    compared_columns = [
        item
        for item in aligned
        if item.old_index is not None
        and item.new_index is not None
        and item.key not in skipped_columns
    ]

    for identity in identities:
        old_row = old_rows.get(identity)
        new_row = new_rows.get(identity)
        key = identity[0]
        if old_row is None and new_row is not None:
            changes.append(RowChange(new_sheet.name, key, "added", new_row=new_row))
            continue
        if old_row is not None and new_row is None:
            changes.append(RowChange(old_sheet.name, key, "deleted", old_row=old_row))
            continue
        if old_row is None or new_row is None:
            continue

        cell_changes: list[CellChange] = []
        for column in compared_columns:
            old_index = column.old_index
            new_index = column.new_index
            if old_index is None or new_index is None:
                continue
            old_cell = old_row.cells[old_index]
            new_cell = new_row.cells[new_index]
            if old_cell == new_cell:
                continue
            header = column.new_header or column.old_header or Header()
            cell_changes.append(
                CellChange(
                    header.display_name(),
                    old_cell,
                    new_cell,
                )
            )
        if cell_changes:
            changes.append(
                RowChange(
                    new_sheet.name,
                    key,
                    "modified",
                    old_row=old_row,
                    new_row=new_row,
                    cell_changes=cell_changes,
                )
            )
    return changes


def _row_identity_map(sheet: TextSheet) -> dict[tuple[str, int], TextRow]:
    return {(row.key, row.occurrence): row for row in sheet.rows}


def format_text_diff(result: DiffResult) -> str:
    """Render a stable, git-like text diff suitable for humans and AI."""
    old_label = Path(result.old_path).name if result.old_path else "old"
    new_label = Path(result.new_path).name if result.new_path else "new"
    lines: list[str] = []
    for sheet in result.sheets:
        if not sheet.has_changes:
            continue
        lines.append(f"diff --table {sheet.sheet}")
        lines.append(f"--- {old_label}")
        lines.append(f"+++ {new_label}")
        if sheet.column_changes:
            lines.append("@@ columns @@")
            for change in sheet.column_changes:
                if change.kind == "added":
                    lines.append(
                        f"+ column {change.column} [{_header_summary(change.new)}]"
                    )
                elif change.kind == "deleted":
                    lines.append(
                        f"- column {change.column} [{_header_summary(change.old)}]"
                    )
                elif change.kind == "modified":
                    lines.append(
                        f"~ column {change.old.display_name() if change.old else change.column}"
                        f" -> {change.new.display_name() if change.new else change.column}"
                    )
                elif change.kind == "moved":
                    lines.append(f"~ column moved {change.column}")
        for change in sheet.row_changes:
            if change.kind == "added":
                lines.append(f"@@ row {change.row} @@")
                lines.append(f"+ row {change.row}")
                for column, value in _row_display_values(change.new_row, sheet.columns):
                    lines.append(f"+   {column}: {value}")
            elif change.kind == "deleted":
                lines.append(f"@@ row {change.row} @@")
                lines.append(f"- row {change.row}")
                for column, value in _row_display_values(change.old_row, sheet.columns):
                    lines.append(f"-   {column}: {value}")
            else:
                lines.append(f"@@ row {change.row} @@")
                for cell_change in change.cell_changes:
                    lines.append(f"- {cell_change.column}: {cell_change.old}")
                    lines.append(f"+ {cell_change.column}: {cell_change.new}")
        lines.append("")
    if not lines:
        lines.append("No differences.")
    if result.issues:
        lines.append("Warnings:")
        for issue in result.issues:
            location = " / ".join(
                str(item)
                for item in (
                    issue.sheet,
                    issue.row,
                    issue.column,
                )
                if item is not None
            )
            lines.append(
                f"- [{issue.severity}] {issue.code}: {issue.message}"
                + (f" ({location})" if location else "")
            )
    return "\n".join(lines)


def _header_summary(header: Header | None) -> str:
    if header is None:
        return ""
    return " | ".join(
        part for part in (header.cn, header.en, header.type) if part
    )


def _row_display_values(
    row: TextRow | None,
    columns: tuple[Header, ...],
) -> list[tuple[str, str]]:
    if row is None:
        return []
    values: list[tuple[str, str]] = []
    for index, cell in enumerate(row.cells):
        if cell == "":
            continue
        header = columns[index] if index < len(columns) else Header()
        values.append((header.display_name(), cell))
    return values


def write_diff_excel(result: DiffResult, output_path: str | Path) -> str:
    """Write a readable change workbook with summary, changes, and issues."""
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    changes = wb.create_sheet("Changes")
    issues = wb.create_sheet("Issues")

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(color="FFFFFF", bold=True)
    added_fill = PatternFill("solid", fgColor="E2F0D9")
    deleted_fill = PatternFill("solid", fgColor="FCE4D6")
    modified_fill = PatternFill("solid", fgColor="FFF2CC")

    summary.append(["Metric", "Value"])
    summary.append(["Old", result.old_path or ""])
    summary.append(["New", result.new_path or ""])
    summary.append(["Comparison", "Text snapshot; formulas use cached values"])
    summary.append(["Has changes", result.has_changes])
    for key, value in result.summary.items():
        summary.append([key, value])
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = header_font
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 72

    changes.append(["Action", "Sheet", "Row", "Column", "Old", "New", "Note"])
    for sheet in result.sheets:
        for change in sheet.column_changes:
            action = {
                "added": "COLUMN_ADDED",
                "deleted": "COLUMN_DELETED",
                "modified": "COLUMN_MODIFIED",
                "moved": "COLUMN_MOVED",
            }[change.kind]
            old_header = change.old.display_name() if change.old else ""
            new_header = change.new.display_name() if change.new else ""
            changes.append(
                [
                    action,
                    sheet.sheet,
                    "",
                    change.column,
                    old_header,
                    new_header,
                    "",
                ]
            )
        for change in sheet.row_changes:
            if change.kind == "added":
                values = " | ".join(
                    f"{name}={value}"
                    for name, value in _row_display_values(change.new_row, sheet.columns)
                )
                changes.append(["ROW_ADDED", sheet.sheet, change.row, "", "", values, ""])
            elif change.kind == "deleted":
                values = " | ".join(
                    f"{name}={value}"
                    for name, value in _row_display_values(change.old_row, sheet.columns)
                )
                changes.append(["ROW_DELETED", sheet.sheet, change.row, "", values, "", ""])
            else:
                for cell_change in change.cell_changes:
                    changes.append(
                        [
                            "CELL_MODIFIED",
                            sheet.sheet,
                            change.row,
                            cell_change.column,
                            cell_change.old,
                            cell_change.new,
                            "",
                        ]
                    )

    for row_cells in changes.iter_rows(min_row=2):
        action = str(row_cells[0].value or "")
        fill = (
            added_fill
            if "ADDED" in action
            else deleted_fill
            if "DELETED" in action
            else modified_fill
        )
        for cell in row_cells:
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in changes[1]:
        cell.fill = header_fill
        cell.font = header_font
    changes.freeze_panes = "A2"
    changes.auto_filter.ref = changes.dimensions
    widths = {"A": 18, "B": 18, "C": 18, "D": 24, "E": 42, "F": 42, "G": 28}
    for column, width in widths.items():
        changes.column_dimensions[column].width = width

    issues.append(["Severity", "Code", "Sheet", "Row", "Column", "Message", "Reference"])
    for issue in result.issues:
        issues.append(
            [
                issue.severity,
                issue.code,
                issue.sheet or "",
                issue.row or "",
                issue.column or "",
                issue.message,
                str(issue.reference or ""),
            ]
        )
    for cell in issues[1]:
        cell.fill = header_fill
        cell.font = header_font
    for column, width in {
        "A": 12,
        "B": 28,
        "C": 18,
        "D": 10,
        "E": 24,
        "F": 70,
        "G": 40,
    }.items():
        issues.column_dimensions[column].width = width
    issues.freeze_panes = "A2"
    issues.auto_filter.ref = issues.dimensions

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    return str(output_path)
