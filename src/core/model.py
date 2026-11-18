"""Pure in-memory workbook model for configuration tables."""

from __future__ import annotations

import os
import shutil
import tempfile
from collections import Counter, OrderedDict
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from .formula import (
    analyze_formula,
    bind_row_local_formula,
    formula_references_sheet_column,
    formula_signature,
)


DATA_START_ROW = 4
_PLACEHOLDER_HEADERS = {"", "none", "null", "nan", "-"}
_ERROR_VALUES = {"#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!"}
_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_HEADER_FONT = Font(color="FFFFFF", bold=True)


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def nonempty(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


@dataclass(slots=True)
class Issue:
    code: str
    severity: str
    message: str
    sheet: str | None = None
    row: int | None = None
    column: str | None = None
    reference: str | None = None

    @property
    def is_error(self) -> bool:
        return self.severity == "error"

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "severity": self.severity,
            "message": self.message,
            "sheet": self.sheet,
            "row": self.row,
            "column": self.column,
            "reference": self.reference,
        }


@dataclass(slots=True)
class Header:
    cn: str = ""
    en: str = ""
    type: str = ""

    @classmethod
    def from_row_values(cls, cn: Any, en: Any, type_value: Any) -> "Header":
        return cls(text(cn), text(en), text(type_value))

    def normalized(self) -> tuple[str, str, str]:
        return (self.cn.strip(), self.en.strip(), self.type.strip())

    def identity(self) -> tuple[str, str, str]:
        return tuple(value.casefold() for value in self.normalized())

    def base_key(self) -> str:
        cn, en, type_value = self.normalized()
        if en and en.casefold() not in _PLACEHOLDER_HEADERS:
            return "en:" + en.casefold()
        return "header:" + "\x1f".join((cn.casefold(), en.casefold(), type_value.casefold()))

    def display_name(self) -> str:
        cn, en, _ = self.normalized()
        if en and cn and en != cn:
            return f"{en} ({cn})"
        return en or cn or "(unnamed)"

    def is_blank(self) -> bool:
        return not any(self.normalized())

    def to_dict(self) -> dict[str, str]:
        return {"cn": self.cn, "en": self.en, "type": self.type}

    @classmethod
    def from_dict(cls, value: Mapping[str, Any]) -> "Header":
        return cls(
            cn=text(value.get("cn", value.get("name_cn", ""))),
            en=text(value.get("en", value.get("name_en", ""))),
            type=text(value.get("type", "")),
        )


@dataclass(slots=True)
class Cell:
    value: Any = None
    formula: str | None = None
    is_array: bool = False
    formula_row: int | None = None
    formula_col: int | None = None

    def is_empty(self) -> bool:
        return not self.formula and not nonempty(self.value)

    def comparison_key(
        self,
        *,
        source_row: int,
        source_col: int,
        include_formula: bool = False,
    ) -> tuple[Any, ...]:
        if include_formula and self.formula:
            return (
                "formula",
                formula_signature(self.formula, source_row, source_col),
                "value",
                self.value if nonempty(self.value) else None,
            )
        if nonempty(self.value):
            return ("value", self.value)
        if self.formula:
            return ("formula", formula_signature(self.formula, source_row, source_col))
        return ("empty",)

    def display(self, *, include_formula: bool = False) -> str:
        if self.formula and include_formula:
            if nonempty(self.value):
                return f"{self.formula} -> {self.value}"
            return self.formula
        if nonempty(self.value):
            return str(self.value)
        return self.formula or ""

    def clone(self) -> "Cell":
        return replace(self)


def cells_equal(
    left: Cell,
    right: Cell,
    *,
    left_row: int,
    left_col: int,
    right_row: int,
    right_col: int,
    include_formula: bool,
) -> bool:
    """Fast cell comparison with formula normalization only as a fallback."""
    left_value = left.value if nonempty(left.value) else None
    right_value = right.value if nonempty(right.value) else None
    if not include_formula and (left_value is not None or right_value is not None):
        return left_value == right_value
    if (
        left.formula == right.formula
        and left.formula is not None
        and left_row == right_row
        and left_col == right_col
        and left_value == right_value
    ):
        return True
    if (
        left.formula is None
        and right.formula is None
        and left_row == right_row
        and left_col == right_col
        and left_value == right_value
    ):
        return True
    return left.comparison_key(
        source_row=left_row,
        source_col=left_col,
        include_formula=include_formula,
    ) == right.comparison_key(
        source_row=right_row,
        source_col=right_col,
        include_formula=include_formula,
    )


@dataclass(slots=True)
class Row:
    key: str
    cells: list[Cell] = field(default_factory=list)
    source_row: int | None = None

    def clone(self) -> "Row":
        return Row(self.key, [cell.clone() for cell in self.cells], self.source_row)

    def pad(self, width: int) -> None:
        if len(self.cells) < width:
            self.cells.extend(Cell() for _ in range(width - len(self.cells)))
        elif len(self.cells) > width:
            del self.cells[width:]

    def is_blank(self) -> bool:
        return not nonempty(self.key) and all(cell.is_empty() for cell in self.cells)


@dataclass(slots=True)
class ValidationReport:
    issues: list[Issue] = field(default_factory=list)
    files: int = 1
    sheets: int = 1
    rows: int = 0
    columns: int = 0
    formulas: int = 0
    dropped_blank_rows: int = 0
    dropped_blank_columns: int = 0

    @property
    def errors(self) -> list[Issue]:
        return [issue for issue in self.issues if issue.is_error]

    @property
    def warnings(self) -> list[Issue]:
        return [issue for issue in self.issues if not issue.is_error]

    @property
    def ok(self) -> bool:
        return not self.errors

    def extend(self, issues: Iterable[Issue]) -> None:
        self.issues.extend(issues)

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "summary": {
                "files": self.files,
                "sheets": self.sheets,
                "rows": self.rows,
                "columns": self.columns,
                "formulas": self.formulas,
                "errors": len(self.errors),
                "warnings": len(self.warnings),
                "dropped_blank_rows": self.dropped_blank_rows,
                "dropped_blank_columns": self.dropped_blank_columns,
            },
            "issues": [issue.to_dict() for issue in self.issues],
        }


class Sheet:
    """One data sheet with three header rows and a preserved row order."""

    def __init__(
        self,
        name: str,
        headers: Sequence[Header] | None = None,
        rows: Sequence[Row] | None = None,
    ) -> None:
        self.name = name
        self.headers = list(headers or [])
        self.rows = list(rows or [])
        self._normalize()

    def _normalize(self) -> None:
        width = len(self.headers)
        for row in self.rows:
            row.pad(width)

    def column_keys(self) -> list[str]:
        counts: Counter[str] = Counter()
        keys: list[str] = []
        for header in self.headers:
            base = header.base_key()
            counts[base] += 1
            occurrence = counts[base]
            keys.append(base if occurrence == 1 else f"{base}#{occurrence}")
        return keys

    def key_to_index(self) -> dict[str, int]:
        return {key: index for index, key in enumerate(self.column_keys())}

    def resolve_column(self, column: str | int) -> int:
        """Resolve a 0-based column index from index, field name, or CN name."""
        if isinstance(column, int):
            index = column
            if 0 <= index < len(self.headers):
                return index
            raise KeyError(f"列索引越界: {column}")

        query = text(column)
        keys = self.column_keys()
        if query in keys:
            return keys.index(query)
        for index, header in enumerate(self.headers):
            if query in {header.en, header.cn}:
                return index
        for index, header in enumerate(self.headers):
            if query.casefold() in {header.en.casefold(), header.cn.casefold()}:
                return index
        raise KeyError(f"列不存在: {column}")

    def row_identities(self) -> list[tuple[str, int]]:
        counts: Counter[str] = Counter()
        identities: list[tuple[str, int]] = []
        for row in self.rows:
            counts[row.key] += 1
            identities.append((row.key, counts[row.key]))
        return identities

    def find_rows(self, key: str) -> list[Row]:
        return [row for row in self.rows if row.key == key]

    def get_row(self, key: str, occurrence: int = 1) -> Row:
        matches = self.find_rows(key)
        if occurrence < 1 or occurrence > len(matches):
            raise KeyError(f"行不存在或出现次数越界: {key}#{occurrence}")
        return matches[occurrence - 1]

    def get_cell(
        self,
        row_key: str,
        column: str | int,
        *,
        occurrence: int = 1,
    ) -> Cell:
        row = self.get_row(row_key, occurrence)
        return row.cells[self.resolve_column(column)]

    def set_cell(
        self,
        row_key: str,
        column: str | int,
        value: Any,
        *,
        formula: str | None = None,
        occurrence: int = 1,
    ) -> Cell:
        row = self.get_row(row_key, occurrence)
        index = self.resolve_column(column)
        formula_text = formula if formula is not None else (
            value if isinstance(value, str) and value.startswith("=") else None
        )
        row.cells[index] = Cell(
            value=None if formula_text else value,
            formula=formula_text,
            formula_row=row.source_row,
            formula_col=index + 1,
        )
        return row.cells[index]

    def add_row(
        self,
        key: str,
        values: Mapping[str | int, Any] | Sequence[Any] | None = None,
        *,
        position: int | None = None,
        source_row: int | None = None,
    ) -> Row:
        width = len(self.headers)
        inferred_row = DATA_START_ROW + (
            len(self.rows) if position is None else max(0, position)
        )
        effective_source_row = source_row or inferred_row
        cells = [Cell() for _ in range(width)]
        if isinstance(values, Mapping):
            for column, value in values.items():
                index = self.resolve_column(column)
                cells[index] = (
                    value.clone()
                    if isinstance(value, Cell)
                    else Cell(
                        value=None if isinstance(value, str) and value.startswith("=") else value,
                        formula=value if isinstance(value, str) and value.startswith("=") else None,
                        formula_row=effective_source_row,
                        formula_col=index + 1,
                    )
                )
        elif values is not None:
            for index, value in enumerate(values[:width]):
                cells[index] = (
                    value.clone()
                    if isinstance(value, Cell)
                    else Cell(
                        value=None if isinstance(value, str) and value.startswith("=") else value,
                        formula=value if isinstance(value, str) and value.startswith("=") else None,
                        formula_row=effective_source_row,
                        formula_col=index + 1,
                    )
                )
        row = Row(str(key), cells, effective_source_row)
        if position is None:
            self.rows.append(row)
        else:
            self.rows.insert(position, row)
        return row

    def delete_row(self, key: str, *, occurrence: int = 1) -> Row:
        row = self.get_row(key, occurrence)
        self.rows.remove(row)
        return row

    def add_column(
        self,
        header: Header,
        *,
        position: int | None = None,
        default: Any = None,
    ) -> None:
        index = len(self.headers) if position is None else position
        if index < 0 or index > len(self.headers):
            raise IndexError(f"列插入位置越界: {index}")
        self.headers.insert(index, header)
        for row in self.rows:
            row.cells.insert(index, Cell(value=default))

    def update_column(
        self,
        column: str | int,
        *,
        cn: str | None = None,
        en: str | None = None,
        type_value: str | None = None,
    ) -> Header:
        index = self.resolve_column(column)
        header = self.headers[index]
        self.headers[index] = Header(
            cn=header.cn if cn is None else cn,
            en=header.en if en is None else en,
            type=header.type if type_value is None else type_value,
        )
        return self.headers[index]

    def delete_column(
        self,
        column: str | int,
        *,
        allow_formula_break: bool = False,
    ) -> Header:
        index = self.resolve_column(column)
        referenced = [
            row.key
            for row in self.rows
            for cell_index, cell in enumerate(row.cells)
            if cell_index != index
            if formula_references_sheet_column(
                cell.formula,
                sheet_name=self.name,
                column_index=index + 1,
            )
        ]
        if referenced and not allow_formula_break:
            sample = ", ".join(referenced[:3])
            raise ValueError(
                f"删除列 {self.headers[index].display_name()} 会破坏公式引用；"
                f"涉及行: {sample}。如确认，请设置 allow_formula_break=True。"
            )
        removed = self.headers.pop(index)
        for row in self.rows:
            if index < len(row.cells):
                row.cells.pop(index)
        return removed

    def remove_empty_rows(self) -> list[int]:
        removed: list[int] = []
        kept: list[Row] = []
        for row in self.rows:
            if row.is_blank():
                if row.source_row is not None:
                    removed.append(row.source_row)
            else:
                kept.append(row)
        self.rows = kept
        return removed

    def remove_empty_columns(self) -> list[str]:
        remove_indexes: list[int] = []
        for index, header in enumerate(self.headers):
            if not header.is_blank():
                continue
            if all(index >= len(row.cells) or row.cells[index].is_empty() for row in self.rows):
                remove_indexes.append(index)
        removed = [self.headers[index].display_name() for index in remove_indexes]
        for index in reversed(remove_indexes):
            del self.headers[index]
            for row in self.rows:
                if index < len(row.cells):
                    del row.cells[index]
        self._normalize()
        return removed

    def clean(self) -> tuple[list[int], list[str]]:
        return self.remove_empty_rows(), self.remove_empty_columns()

class WorkbookDocument:
    """A dependency-light representation of an xlsx configuration workbook."""

    def __init__(
        self,
        sheets: Mapping[str, Sheet] | None = None,
        *,
        source_path: str | Path | None = None,
    ) -> None:
        self.sheets: OrderedDict[str, Sheet] = OrderedDict(sheets or {})
        self.source_path = str(source_path) if source_path is not None else None
        self.load_issues: list[Issue] = []

    @classmethod
    def from_xlsx(
        cls,
        path: str | Path,
        *,
        data_sheets_only: bool = True,
        include_hidden_sheets: bool = False,
        include_values: bool = True,
    ) -> "WorkbookDocument":
        path = Path(path)
        actual_path, temp_path = _readable_xlsx_path(path)
        wb_values = load_workbook(actual_path, data_only=True, read_only=True)
        wb_formulas = load_workbook(actual_path, data_only=False, read_only=True)
        document = cls(source_path=path)
        try:
            selected = [
                name
                for name in wb_values.sheetnames
                if (not data_sheets_only or name.startswith("#"))
            ]
            for sheet_name in selected:
                if sheet_name not in wb_formulas.sheetnames:
                    document.load_issues.append(
                        Issue(
                            "sheet_missing_formula_read",
                            "warning",
                            "该 sheet 无法读取公式层，将仅使用缓存值。",
                            sheet=sheet_name,
                        )
                    )
                    formula_rows: list[tuple[Any, ...]] = []
                else:
                    formula_rows = [
                        tuple(row)
                        for row in wb_formulas[sheet_name].iter_rows(values_only=True)
                    ]
                value_rows = [
                    tuple(row) for row in wb_values[sheet_name].iter_rows(values_only=True)
                ]
                sheet = _sheet_from_rows(
                    sheet_name,
                    value_rows,
                    formula_rows,
                    document.load_issues,
                    include_values=include_values,
                )
                document.sheets[sheet_name] = sheet
        finally:
            wb_values.close()
            wb_formulas.close()
            if temp_path is not None:
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass
        return document

    @classmethod
    def from_sheets(
        cls,
        sheets: Mapping[str, Sheet],
        *,
        source_path: str | Path | None = None,
    ) -> "WorkbookDocument":
        return cls(OrderedDict((name, sheet) for name, sheet in sheets.items()), source_path=source_path)

    def sheet(self, name: str) -> Sheet:
        try:
            return self.sheets[name]
        except KeyError as exc:
            raise KeyError(f"sheet 不存在: {name}") from exc

    def add_sheet(self, sheet: Sheet, *, position: int | None = None) -> Sheet:
        if sheet.name in self.sheets:
            raise ValueError(f"sheet 已存在: {sheet.name}")
        if position is None:
            self.sheets[sheet.name] = sheet
        else:
            items = list(self.sheets.items())
            items.insert(position, (sheet.name, sheet))
            self.sheets = OrderedDict(items)
        return sheet

    def delete_sheet(self, name: str) -> Sheet:
        return self.sheets.pop(name)

    def clone(self) -> "WorkbookDocument":
        return WorkbookDocument(
            OrderedDict(
                (
                    name,
                    Sheet(
                        sheet.name,
                        [replace(header) for header in sheet.headers],
                        [row.clone() for row in sheet.rows],
                    ),
                )
                for name, sheet in self.sheets.items()
            ),
            source_path=self.source_path,
        )

    def clean(self) -> tuple[list[int], list[str]]:
        rows: list[int] = []
        columns: list[str] = []
        for sheet in self.sheets.values():
            removed_rows, removed_columns = sheet.clean()
            rows.extend(removed_rows)
            columns.extend(removed_columns)
        return rows, columns

    def validate(self, *, clean_first: bool = False) -> ValidationReport:
        report = ValidationReport(files=1, sheets=len(self.sheets))
        report.extend(self.load_issues)
        formula_groups: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        dropped_rows: list[int] = []
        dropped_cols: list[str] = []
        if clean_first:
            dropped_rows, dropped_cols = self.clean()
        report.dropped_blank_rows = len(dropped_rows)
        report.dropped_blank_columns = len(dropped_cols)

        for sheet in self.sheets.values():
            report.rows += len(sheet.rows)
            report.columns += len(sheet.headers)
            report.extend(_validate_headers(sheet))
            report.extend(_validate_rows(sheet))
            for row in sheet.rows:
                for index, cell in enumerate(row.cells):
                    column = sheet.headers[index] if index < len(sheet.headers) else Header()
                    if cell.formula:
                        report.formulas += 1
                        findings = analyze_formula(
                            cell.formula,
                            source_row=cell.formula_row or row.source_row or DATA_START_ROW,
                            source_col=cell.formula_col or index + 1,
                            is_array=cell.is_array,
                        )
                        for finding in findings:
                            key = (
                                sheet.name,
                                column.display_name(),
                                finding.code,
                                finding.severity,
                            )
                            group = formula_groups.get(key)
                            if group is None:
                                formula_groups[key] = {
                                    "count": 1,
                                    "row": cell.formula_row,
                                    "message": finding.message,
                                    "reference": finding.reference or cell.formula,
                                }
                            else:
                                group["count"] += 1
                    if text(cell.value) in _ERROR_VALUES:
                        report.issues.append(
                            Issue(
                                "excel_error_value",
                                "error",
                                f"检测到 Excel 错误值: {text(cell.value)}",
                                sheet=sheet.name,
                                row=row.source_row,
                                column=column.display_name(),
                            )
                        )
        for (sheet_name, column_name, code, severity), group in formula_groups.items():
            count = group["count"]
            message = group["message"]
            if count > 1:
                message += f" 共 {count} 个单元格触发。"
            report.issues.append(
                Issue(
                    code,
                    severity,
                    message,
                    sheet=sheet_name,
                    row=group["row"],
                    column=column_name,
                    reference=group["reference"],
                )
            )
        return report

    def save(self, path: str | Path) -> str:
        """Write a clean output workbook.

        This writer is intended for generated outputs and simple-formula
        configuration tables. It does not preserve Excel COM caches, macros,
        external-link XML, or dynamic-array metadata.
        """
        self.clean()
        wb = Workbook()
        wb.remove(wb.active)
        if not self.sheets:
            ws = wb.create_sheet("NoData")
            ws["A1"] = "No data sheets"
        for sheet in self.sheets.values():
            _write_sheet(wb, sheet)
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)
        wb.close()
        return str(path)


def _readable_xlsx_path(path: Path) -> tuple[str, str | None]:
    if path.suffix.casefold() in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return str(path), None
    handle = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    handle.close()
    shutil.copy2(path, handle.name)
    return handle.name, handle.name


def _formula_text(value: Any) -> tuple[str | None, bool]:
    if isinstance(value, ArrayFormula):
        return value.text, True
    if isinstance(value, str) and value.startswith("="):
        return value, False
    return None, False


def _cell_from_rows(value: Any, formula_value: Any, *, include_values: bool) -> Cell:
    formula, is_array = _formula_text(formula_value)
    return Cell(
        value=value if include_values else None,
        formula=formula,
        is_array=is_array,
    )


def _sheet_from_rows(
    name: str,
    value_rows: Sequence[Sequence[Any]],
    formula_rows: Sequence[Sequence[Any]],
    issues: list[Issue],
    *,
    include_values: bool,
) -> Sheet:
    row_count = max(len(value_rows), len(formula_rows))
    width = max(
        max((len(row) for row in value_rows), default=0),
        max((len(row) for row in formula_rows), default=0),
    )

    def at(rows: Sequence[Sequence[Any]], row_index: int, col_index: int) -> Any:
        if row_index >= len(rows):
            return None
        row = rows[row_index]
        return row[col_index] if col_index < len(row) else None

    used_columns: list[int] = []
    for col_index in range(width):
        used = False
        for row_index in range(row_count):
            value = at(value_rows, row_index, col_index)
            formula, _ = _formula_text(at(formula_rows, row_index, col_index))
            if nonempty(value) or formula:
                used = True
                break
        if used:
            used_columns.append(col_index)

    headers = [
        Header.from_row_values(
            at(value_rows, 0, col_index),
            at(value_rows, 1, col_index),
            at(value_rows, 2, col_index),
        )
        for col_index in used_columns
    ]

    rows: list[Row] = []
    for row_index in range(DATA_START_ROW - 1, row_count):
        cells = [
            _cell_from_rows(
                at(value_rows, row_index, col_index),
                at(formula_rows, row_index, col_index),
                include_values=include_values,
            )
            for col_index in used_columns
        ]
        if all(cell.is_empty() for cell in cells):
            if row_index + 1 < row_count:
                issues.append(
                    Issue(
                        "blank_row_removed",
                        "warning",
                        "检测到空数据行，加载时已从模型中移除。",
                        sheet=name,
                        row=row_index + 1,
                    )
                )
            continue

        key_value = cells[0].value if cells else None
        key_from_formula = False
        if not nonempty(key_value) and cells and cells[0].formula:
            if len(cells) > 1 and nonempty(cells[1].value):
                key_value = cells[1].value
            else:
                key_value = cells[0].formula
            key_from_formula = True
        if not nonempty(key_value):
            key_value = f"__row_{row_index + 1}"
            issues.append(
                Issue(
                    "missing_row_key",
                    "warning",
                    "数据行没有 RowName；已使用临时行键，保存/合并前应补全。",
                    sheet=name,
                    row=row_index + 1,
                )
            )
        elif key_from_formula:
            issues.append(
                Issue(
                    "row_key_from_formula",
                    "warning",
                    "RowName 来自公式且缺少缓存值；当前使用计算/公式后备值作为行键。",
                    sheet=name,
                    row=row_index + 1,
                )
            )

        for col_index, cell in enumerate(cells):
            cell.formula_row = row_index + 1 if cell.formula else None
            cell.formula_col = used_columns[col_index] + 1 if cell.formula else None
        rows.append(Row(str(key_value).strip(), cells, row_index + 1))

    return Sheet(name, headers, rows)


def _validate_headers(sheet: Sheet) -> list[Issue]:
    issues: list[Issue] = []
    cn_map: dict[str, list[int]] = {}
    en_map: dict[str, list[int]] = {}
    triple_map: dict[tuple[str, str, str], list[int]] = {}

    for index, header in enumerate(sheet.headers):
        cn, en, type_value = header.normalized()
        if cn:
            cn_map.setdefault(cn.casefold(), []).append(index)
        if en and en.casefold() not in _PLACEHOLDER_HEADERS:
            en_map.setdefault(en.casefold(), []).append(index)
        if cn or en or type_value:
            triple_map.setdefault(header.identity(), []).append(index)
        if not en and any(
            index < len(row.cells) and not row.cells[index].is_empty()
            for row in sheet.rows
        ):
            issues.append(
                Issue(
                    "missing_english_header",
                    "warning",
                    "该列有数据，但缺少英文字段名。",
                    sheet=sheet.name,
                    column=get_column_letter(index + 1),
                )
            )

    for label, mapping, code, severity in (
        ("英文字段名", en_map, "duplicate_english_header", "error"),
        ("中文字段名", cn_map, "duplicate_chinese_header", "warning"),
        ("完整表头", triple_map, "duplicate_header", "warning"),
    ):
        for value, indexes in mapping.items():
            if len(indexes) <= 1:
                continue
            issues.append(
                Issue(
                    code,
                    severity,
                    f"{label}重复；列: "
                    + ", ".join(get_column_letter(index + 1) for index in indexes),
                    sheet=sheet.name,
                    column=get_column_letter(indexes[0] + 1),
                    reference=" | ".join(value) if isinstance(value, tuple) else value,
                )
            )
    return issues


def _validate_rows(sheet: Sheet) -> list[Issue]:
    issues: list[Issue] = []
    key_map: dict[str, list[int]] = {}
    for row in sheet.rows:
        key_map.setdefault(row.key, []).append(row.source_row or 0)
    for key, row_numbers in key_map.items():
        if len(row_numbers) <= 1:
            continue
        issues.append(
            Issue(
                "duplicate_row_key",
                "error",
                "重复 RowName；所有重复行都保留在模型中，但三路合并不会静默覆盖。",
                sheet=sheet.name,
                row=row_numbers[0],
                reference=key,
            )
        )
    return issues


def _write_sheet(wb: Workbook, sheet: Sheet) -> None:
    sheet.clean()
    ws = wb.create_sheet(sheet.name)
    for col_index, header in enumerate(sheet.headers, start=1):
        for row_index, value in enumerate(
            (header.cn, header.en, header.type),
            start=1,
        ):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            cell.fill = _HEADER_FILL
            cell.font = _HEADER_FONT
        ws.column_dimensions[get_column_letter(col_index)].width = 18

    for output_offset, row in enumerate(sheet.rows):
        excel_row = DATA_START_ROW + output_offset
        ws.cell(row=excel_row, column=1, value=row.key)
        for col_index, cell in enumerate(row.cells, start=1):
            if cell.formula:
                formula = bind_row_local_formula(
                    cell.formula,
                    source_row=cell.formula_row or excel_row,
                    source_col=cell.formula_col or col_index,
                    target_row=excel_row,
                    target_col=col_index,
                )
                if cell.is_array:
                    ws.cell(row=excel_row, column=col_index).value = ArrayFormula(
                        f"{get_column_letter(col_index)}{excel_row}",
                        formula,
                    )
                else:
                    ws.cell(row=excel_row, column=col_index, value=formula)
            elif nonempty(cell.value):
                ws.cell(row=excel_row, column=col_index, value=cell.value)
    ws.freeze_panes = f"A{DATA_START_ROW}"
    ws.auto_filter.ref = (
        f"A3:{get_column_letter(max(1, len(sheet.headers)))}"
        f"{max(DATA_START_ROW - 1, DATA_START_ROW + len(sheet.rows) - 1)}"
    )
