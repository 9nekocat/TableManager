"""Small CRUD operation layer shared by humans and AI callers."""

from __future__ import annotations

from typing import Any, Iterable, Mapping

from .model import Header, WorkbookDocument


def apply_operations(
    document: WorkbookDocument,
    operations: Iterable[Mapping[str, Any]],
) -> WorkbookDocument:
    """Apply declarative cell/row/column operations in order."""
    for index, operation in enumerate(operations, start=1):
        op = str(operation.get("op", "")).strip().lower()
        sheet_name = str(operation.get("sheet", ""))
        try:
            if op == "set_cell":
                sheet = document.sheet(sheet_name)
                sheet.set_cell(
                    str(operation["row"]),
                    operation["column"],
                    operation.get("value"),
                    formula=operation.get("formula"),
                    occurrence=int(operation.get("occurrence", 1)),
                )
            elif op == "add_row":
                sheet = document.sheet(sheet_name)
                key = str(operation.get("row", operation.get("key", "")))
                if not key:
                    raise ValueError("add_row 需要 row 或 key")
                sheet.add_row(
                    key,
                    operation.get("values"),
                    position=_optional_int(operation.get("position")),
                )
            elif op == "delete_row":
                sheet = document.sheet(sheet_name)
                sheet.delete_row(
                    str(operation["row"]),
                    occurrence=int(operation.get("occurrence", 1)),
                )
            elif op == "add_column":
                sheet = document.sheet(sheet_name)
                sheet.add_column(
                    Header.from_dict(operation["header"]),
                    position=_optional_int(operation.get("position")),
                    default=operation.get("default"),
                )
            elif op == "update_column":
                sheet = document.sheet(sheet_name)
                sheet.update_column(
                    operation["column"],
                    cn=operation.get("cn"),
                    en=operation.get("en"),
                    type_value=operation.get("type"),
                )
            elif op == "delete_column":
                sheet = document.sheet(sheet_name)
                sheet.delete_column(
                    operation["column"],
                    allow_formula_break=bool(operation.get("allow_formula_break", False)),
                )
            else:
                raise ValueError(f"未知操作: {op or '(empty)'}")
        except Exception as exc:
            raise ValueError(f"第 {index} 个操作失败 ({op}): {exc}") from exc
    return document


def _optional_int(value: Any) -> int | None:
    return None if value is None else int(value)
