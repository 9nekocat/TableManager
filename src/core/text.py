"""Stable text snapshots used by the diff pipeline."""

from __future__ import annotations

from collections import Counter, OrderedDict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any

from .model import Header, Issue, WorkbookDocument, nonempty


@dataclass(frozen=True, slots=True)
class TextColumn:
    key: str
    header: Header


@dataclass(frozen=True, slots=True)
class TextRow:
    key: str
    occurrence: int
    cells: tuple[str, ...]
    source_row: int | None = None


@dataclass(frozen=True, slots=True)
class TextSheet:
    name: str
    columns: tuple[TextColumn, ...]
    rows: tuple[TextRow, ...]

    @property
    def headers(self) -> tuple[Header, ...]:
        return tuple(column.header for column in self.columns)

    @property
    def column_keys(self) -> tuple[str, ...]:
        return tuple(column.key for column in self.columns)


@dataclass(frozen=True, slots=True)
class TextWorkbook:
    source_path: str | None
    sheets: OrderedDict[str, TextSheet]
    issues: tuple[Issue, ...] = ()


def sheet_headers_to_text(sheet) -> TextSheet:
    """Create a header-only text sheet for column alignment."""
    columns = tuple(
        TextColumn(key, header)
        for key, header in zip(sheet.column_keys(), sheet.headers)
    )
    return TextSheet(sheet.name, columns, ())


def value_to_text(value: Any) -> str:
    """Convert a cached Excel value to a deterministic comparison string."""
    if not nonempty(value):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, ".15g")
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    return str(value)


def workbook_to_text(document: WorkbookDocument) -> TextWorkbook:
    """Freeze a workbook to text using cached values only.

    Formula cells intentionally use ``cell.value``. If Excel has no cached
    result, the cell becomes empty and a grouped warning is emitted.
    """
    sheets: OrderedDict[str, TextSheet] = OrderedDict()
    missing_cache: dict[tuple[str, str], dict[str, Any]] = {}
    for sheet in document.sheets.values():
        column_keys = sheet.column_keys()
        columns = tuple(
            TextColumn(key, header)
            for key, header in zip(column_keys, sheet.headers)
        )
        occurrences: Counter[str] = Counter()
        rows: list[TextRow] = []
        for row in sheet.rows:
            occurrences[row.key] += 1
            texts: list[str] = []
            for index, cell in enumerate(row.cells):
                if cell.formula and not nonempty(cell.value):
                    header = sheet.headers[index] if index < len(sheet.headers) else Header()
                    key = (sheet.name, header.display_name())
                    group = missing_cache.get(key)
                    if group is None:
                        missing_cache[key] = {
                            "count": 1,
                            "row": cell.formula_row or row.source_row,
                            "formula": cell.formula,
                        }
                    else:
                        group["count"] += 1
                texts.append(value_to_text(cell.value))
            rows.append(
                TextRow(
                    row.key,
                    occurrences[row.key],
                    tuple(texts),
                    row.source_row,
                )
            )
        sheets[sheet.name] = TextSheet(sheet.name, columns, tuple(rows))

    issues: list[Issue] = []
    for (sheet_name, column_name), group in missing_cache.items():
        count = group["count"]
        issues.append(
            Issue(
                "formula_value_unavailable",
                "warning",
                (
                    "公式没有可用的 Excel 缓存值，差异中按空值处理。"
                    + (f" 共 {count} 个单元格。" if count > 1 else "")
                ),
                sheet=sheet_name,
                row=group["row"],
                column=column_name,
                reference=str(group["formula"])[:120],
            )
        )
    return TextWorkbook(document.source_path, sheets, tuple(issues))


def render_text_workbook(snapshot: TextWorkbook) -> str:
    """Render a snapshot as a deterministic, AI-readable table text."""
    lines: list[str] = []
    for sheet in snapshot.sheets.values():
        lines.append(f"## {sheet.name}")
        lines.append(
            "columns\t"
            + "\t".join(
                f"{column.header.cn}|{column.header.en}|{column.header.type}"
                for column in sheet.columns
            )
        )
        for row in sheet.rows:
            lines.append(
                f"row\t{row.key}#{row.occurrence}\t"
                + "\t".join(cell.replace("\t", "\\t").replace("\n", "\\n") for cell in row.cells)
            )
    return "\n".join(lines)
