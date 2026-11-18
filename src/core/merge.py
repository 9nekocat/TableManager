"""Three-way merge for the project's simple-formula configuration tables."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Mapping

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .diff import ColumnChange, _align_columns
from .model import Cell, Header, Issue, Row, Sheet, WorkbookDocument, nonempty, text


VALID_STRATEGIES = {"fail", "ours", "theirs"}


@dataclass(slots=True)
class MergeChange:
    sheet: str
    kind: str
    row: str = ""
    column: str = ""
    base: str = ""
    ours: str = ""
    theirs: str = ""
    result: str = ""
    conflict: bool = False
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "kind": self.kind,
            "row": self.row,
            "column": self.column,
            "base": self.base,
            "ours": self.ours,
            "theirs": self.theirs,
            "result": self.result,
            "conflict": self.conflict,
            "note": self.note,
        }


@dataclass(slots=True)
class MergeResult:
    document: WorkbookDocument
    strategy: str
    changes: list[MergeChange] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)

    @property
    def conflicts(self) -> list[MergeChange]:
        return [change for change in self.changes if change.conflict]

    @property
    def ok(self) -> bool:
        return not self.conflicts and not any(issue.is_error for issue in self.issues)

    def require_resolved(self) -> None:
        if self.conflicts:
            sample = ", ".join(
                f"{change.sheet}!{change.row or '-'}.{change.column}"
                for change in self.conflicts[:5]
            )
            raise MergeConflictError(
                f"存在 {len(self.conflicts)} 个未解决冲突: {sample}"
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "strategy": self.strategy,
            "summary": {
                "changes": len(self.changes),
                "conflicts": len(self.conflicts),
                "errors": len([issue for issue in self.issues if issue.is_error]),
                "warnings": len([issue for issue in self.issues if not issue.is_error]),
            },
            "changes": [change.to_dict() for change in self.changes],
            "issues": [issue.to_dict() for issue in self.issues],
        }


class MergeConflictError(RuntimeError):
    pass


@dataclass(slots=True)
class _ColumnMap:
    headers: list[Header]
    base: dict[int, int]
    ours: dict[int, int]
    theirs: dict[int, int]


def merge_files(
    base_path: str | Path,
    ours_path: str | Path,
    theirs_path: str | Path,
    *,
    strategy: str = "fail",
) -> MergeResult:
    base = WorkbookDocument.from_xlsx(base_path)
    ours = WorkbookDocument.from_xlsx(ours_path)
    theirs = WorkbookDocument.from_xlsx(theirs_path)
    result = merge_documents(base, ours, theirs, strategy=strategy)
    result.document.source_path = str(ours_path)
    return result


def merge_documents(
    base: WorkbookDocument,
    ours: WorkbookDocument,
    theirs: WorkbookDocument,
    *,
    strategy: str = "fail",
) -> MergeResult:
    if strategy not in VALID_STRATEGIES:
        raise ValueError(
            f"未知合并策略: {strategy}; 可选值: {', '.join(sorted(VALID_STRATEGIES))}"
        )
    merged = WorkbookDocument()
    result = MergeResult(merged, strategy)
    result.issues.extend(base.load_issues)
    result.issues.extend(ours.load_issues)
    result.issues.extend(theirs.load_issues)

    sheet_names = list(
        OrderedDict.fromkeys(
            [*base.sheets.keys(), *ours.sheets.keys(), *theirs.sheets.keys()]
        ).keys()
    )
    for sheet_name in sheet_names:
        _merge_sheet(
            sheet_name,
            base.sheets.get(sheet_name),
            ours.sheets.get(sheet_name),
            theirs.sheets.get(sheet_name),
            result,
        )

    report = merged.validate(clean_first=True)
    result.issues.extend(report.issues)
    return result


def _merge_sheet(
    sheet_name: str,
    base: Sheet | None,
    ours: Sheet | None,
    theirs: Sheet | None,
    result: MergeResult,
) -> None:
    if base is None:
        if ours is not None and theirs is None:
            result.document.add_sheet(_copy_sheet(ours))
            result.changes.append(
                MergeChange(sheet_name, "sheet_added", note="ours")
            )
            return
        if ours is None and theirs is not None:
            result.document.add_sheet(_copy_sheet(theirs))
            result.changes.append(
                MergeChange(sheet_name, "sheet_added", note="theirs")
            )
            return
        if ours is None and theirs is None:
            return
    elif ours is None and theirs is None:
        result.changes.append(MergeChange(sheet_name, "sheet_deleted", note="both"))
        return
    elif ours is None or theirs is None:
        present = theirs if ours is None else ours
        side = "theirs" if ours is None else "ours"
        if present is not None and _sheet_equal(base, present):
            result.changes.append(
                MergeChange(sheet_name, "sheet_deleted", note=side)
            )
            return
        chosen = present if strategy_choice(result.strategy, "ours") == "ours" else base
        chosen_side = side if chosen is present else "base"
        result.document.add_sheet(_copy_sheet(chosen)) if chosen is not None else None
        result.changes.append(
            MergeChange(
                sheet_name,
                "sheet_delete_conflict",
                conflict=True,
                result=chosen_side,
                note=f"{side} 删除，但另一侧修改了该 sheet；按 {result.strategy} 保留 {chosen_side}",
            )
        )
        return

    assert base is not None and ours is not None and theirs is not None
    column_map, column_changes = _merge_columns(
        sheet_name,
        base,
        ours,
        theirs,
        result,
    )
    result.changes.extend(column_changes)
    if not column_map.headers:
        result.changes.append(MergeChange(sheet_name, "sheet_deleted", note="all columns deleted"))
        return

    merged_rows = _merge_rows(
        sheet_name,
        base,
        ours,
        theirs,
        column_map,
        result,
    )
    if merged_rows:
        result.document.add_sheet(Sheet(sheet_name, column_map.headers, merged_rows))
    else:
        result.document.add_sheet(Sheet(sheet_name, column_map.headers, []))


def strategy_choice(strategy: str, fallback: str) -> str:
    return fallback if strategy == "fail" else strategy


def _copy_sheet(sheet: Sheet) -> Sheet:
    return Sheet(
        sheet.name,
        [replace(header) for header in sheet.headers],
        [row.clone() for row in sheet.rows],
    )


def _sheet_equal(left: Sheet, right: Sheet) -> bool:
    if [header.identity() for header in left.headers] != [
        header.identity() for header in right.headers
    ]:
        return False
    if len(left.rows) != len(right.rows):
        return False
    for left_row, right_row in zip(left.rows, right.rows):
        if left_row.key != right_row.key:
            return False
        for index, (left_cell, right_cell) in enumerate(zip(left_row.cells, right_row.cells)):
            if left_cell.comparison_key(
                source_row=left_cell.formula_row or left_row.source_row or 4,
                source_col=left_cell.formula_col or index + 1,
                include_formula=True,
            ) != right_cell.comparison_key(
                source_row=right_cell.formula_row or right_row.source_row or 4,
                source_col=right_cell.formula_col or index + 1,
                include_formula=True,
            ):
                return False
    return True


def _base_column_mapping(base: Sheet, branch: Sheet) -> dict[int, tuple[int | None, Header]]:
    aligned, _ = _align_columns(base, branch)
    result: dict[int, tuple[int | None, Header]] = {}
    for item in aligned:
        if item.old_index is None:
            continue
        result[item.old_index] = (
            item.new_index,
            item.new_header or item.old_header or Header(),
        )
    return result


def _column_changed(
    base: Sheet,
    branch: Sheet,
    base_index: int,
    branch_index: int | None,
) -> bool:
    if branch_index is None:
        return True
    if base.headers[base_index].identity() != branch.headers[branch_index].identity():
        return True
    base_rows = _row_identity_map(base)
    branch_rows = _row_identity_map(branch)
    for identity in set(base_rows) | set(branch_rows):
        base_row = base_rows.get(identity)
        branch_row = branch_rows.get(identity)
        if base_row is None or branch_row is None:
            if base_row is not None or branch_row is not None:
                return True
            continue
        base_cell = base_row.cells[base_index]
        branch_cell = branch_row.cells[branch_index]
        if base_cell.comparison_key(
            source_row=base_cell.formula_row or base_row.source_row or 4,
            source_col=base_cell.formula_col or base_index + 1,
            include_formula=True,
        ) != branch_cell.comparison_key(
            source_row=branch_cell.formula_row or branch_row.source_row or 4,
            source_col=branch_cell.formula_col or branch_index + 1,
            include_formula=True,
        ):
            return True
    return False


def _merge_columns(
    sheet_name: str,
    base: Sheet,
    ours: Sheet,
    theirs: Sheet,
    result: MergeResult,
) -> tuple[_ColumnMap, list[MergeChange]]:
    ours_map = _base_column_mapping(base, ours)
    theirs_map = _base_column_mapping(base, theirs)
    output_headers: list[Header] = []
    base_to_output: dict[int, int] = {}
    ours_to_output: dict[int, int] = {}
    theirs_to_output: dict[int, int] = {}
    changes: list[MergeChange] = []

    def add_output(header: Header, base_idx: int | None, ours_idx: int | None, theirs_idx: int | None) -> None:
        output_index = len(output_headers)
        output_headers.append(header)
        if base_idx is not None:
            base_to_output[base_idx] = output_index
        if ours_idx is not None:
            ours_to_output[ours_idx] = output_index
        if theirs_idx is not None:
            theirs_to_output[theirs_idx] = output_index

    for base_index, base_header in enumerate(base.headers):
        ours_index, ours_header = ours_map.get(base_index, (None, base_header))
        theirs_index, theirs_header = theirs_map.get(base_index, (None, base_header))
        ours_changed = _column_changed(base, ours, base_index, ours_index)
        theirs_changed = _column_changed(base, theirs, base_index, theirs_index)

        if ours_index is None and theirs_index is None:
            changes.append(
                MergeChange(
                    sheet_name,
                    "column_deleted",
                    column=base_header.display_name(),
                    base=base_header.display_name(),
                    result="",
                    note="both",
                )
            )
            continue
        if ours_index is None or theirs_index is None:
            missing_side = "ours" if ours_index is None else "theirs"
            kept_side = "theirs" if missing_side == "ours" else "ours"
            kept_changed = theirs_changed if missing_side == "ours" else ours_changed
            if not kept_changed:
                changes.append(
                    MergeChange(
                        sheet_name,
                        "column_deleted",
                        column=base_header.display_name(),
                        base=base_header.display_name(),
                        result="",
                        note=missing_side,
                    )
                )
                continue
            conflict = True
            # A deleted column cannot be the preview value. Keep the only
            # concrete branch and mark the result as a conflict for review.
            chosen = kept_side
            chosen_index = ours_index if chosen == "ours" else theirs_index
            chosen_header = (
                ours.headers[ours_index]
                if chosen_index is not None and chosen == "ours"
                else theirs.headers[chosen_index]
                if chosen_index is not None
                else base_header
            )
            add_output(chosen_header, None, ours_index, theirs_index)
            changes.append(
                MergeChange(
                    sheet_name,
                    "column_delete_conflict",
                    column=base_header.display_name(),
                    base=base_header.display_name(),
            ours="" if ours_index is None else ours_header.display_name(),
            theirs="" if theirs_index is None else theirs_header.display_name(),
                    result=chosen_header.display_name(),
                    conflict=conflict,
                    note=f"{missing_side} 删除，但 {kept_side} 修改了该列",
                )
            )
            continue

        if ours_header.identity() == theirs_header.identity():
            merged_header = replace(ours_header)
        elif ours_header.identity() == base_header.identity():
            merged_header = replace(theirs_header)
        elif theirs_header.identity() == base_header.identity():
            merged_header = replace(ours_header)
        else:
            chosen = strategy_choice(result.strategy, "ours")
            merged_header = replace(ours_header if chosen == "ours" else theirs_header)
            changes.append(
                MergeChange(
                    sheet_name,
                    "column_header_conflict",
                    column=base_header.display_name(),
                    base=base_header.display_name(),
                    ours=ours_header.display_name(),
                    theirs=theirs_header.display_name(),
                    result=merged_header.display_name(),
                    conflict=True,
                    note=f"表头双方均修改；按 {result.strategy} 选择 {chosen}",
                )
            )
        add_output(merged_header, base_index, ours_index, theirs_index)
        if merged_header.identity() != base_header.identity():
            changes.append(
                MergeChange(
                    sheet_name,
                    "column_modified",
                    column=base_header.display_name(),
                    base=base_header.display_name(),
                    ours=ours_header.display_name(),
                    theirs=theirs_header.display_name(),
                    result=merged_header.display_name(),
                )
            )

    # Columns not present in base. Ours order wins, then theirs-only additions.
    mapped_ours = {index for index, _ in ours_map.values() if index is not None}
    mapped_theirs = {index for index, _ in theirs_map.values() if index is not None}
    output_keys = {header.base_key() for header in output_headers}
    for ours_index, header in enumerate(ours.headers):
        if ours_index in mapped_ours:
            continue
        key = header.base_key()
        if key in output_keys:
            continue
        add_output(replace(header), None, ours_index, None)
        output_keys.add(key)
        changes.append(
            MergeChange(
                sheet_name,
                "column_added",
                column=header.display_name(),
                ours=header.display_name(),
                result=header.display_name(),
                note="ours",
            )
        )
    for theirs_index, header in enumerate(theirs.headers):
        if theirs_index in mapped_theirs:
            continue
        key = header.base_key()
        existing_index = next(
            (
                index
                for index, existing in enumerate(output_headers)
                if existing.base_key() == key
            ),
            None,
        )
        if existing_index is not None:
            theirs_to_output[theirs_index] = existing_index
            existing_header = output_headers[existing_index]
            if existing_header.identity() != header.identity():
                chosen = strategy_choice(result.strategy, "ours")
                output_headers[existing_index] = replace(
                    existing_header if chosen == "ours" else header
                )
                changes.append(
                    MergeChange(
                        sheet_name,
                        "column_header_conflict",
                        column=existing_header.display_name(),
                        ours=existing_header.display_name(),
                        theirs=header.display_name(),
                        result=output_headers[existing_index].display_name(),
                        conflict=True,
                        note=f"双方新增同名列但表头不同；按 {result.strategy} 选择 {chosen}",
                    )
                )
            continue
        add_output(replace(header), None, None, theirs_index)
        output_keys.add(key)
        changes.append(
            MergeChange(
                sheet_name,
                "column_added",
                column=header.display_name(),
                theirs=header.display_name(),
                result=header.display_name(),
                note="theirs",
            )
        )

    return (
        _ColumnMap(output_headers, base_to_output, ours_to_output, theirs_to_output),
        changes,
    )


def _row_identity_map(sheet: Sheet) -> dict[tuple[str, int], Row]:
    return dict(zip(sheet.row_identities(), sheet.rows))


def _remap_row(row: Row, source_to_output: Mapping[int, int], width: int) -> Row:
    cells = [Cell() for _ in range(width)]
    for source_index, output_index in source_to_output.items():
        if source_index < len(row.cells) and output_index < width:
            cells[output_index] = row.cells[source_index].clone()
    return Row(row.key, cells, row.source_row)


def _cell_key(cell: Cell, row: Row | None, source_index: int | None) -> tuple[Any, ...]:
    return cell.comparison_key(
        source_row=cell.formula_row or (row.source_row if row else 4) or 4,
        source_col=cell.formula_col or (source_index + 1 if source_index is not None else 1),
        include_formula=True,
    )


def _row_changed(
    base_row: Row,
    branch_row: Row,
    base_sheet: Sheet,
    branch_sheet: Sheet,
    branch_to_base_columns: Mapping[int, int],
) -> bool:
    for branch_index, base_index in branch_to_base_columns.items():
        if branch_index >= len(branch_row.cells) or base_index >= len(base_row.cells):
            continue
        if _cell_key(
            branch_row.cells[branch_index],
            branch_row,
            branch_index,
        ) != _cell_key(
            base_row.cells[base_index],
            base_row,
            base_index,
        ):
            return True
    return False


def _merge_cell(
    sheet_name: str,
    row_key: str,
    column_name: str,
    base_cell: Cell,
    ours_cell: Cell,
    theirs_cell: Cell,
    base_row: Row | None,
    ours_row: Row | None,
    theirs_row: Row | None,
    base_index: int | None,
    ours_index: int | None,
    theirs_index: int | None,
    result: MergeResult,
) -> Cell:
    base_key = _cell_key(base_cell, base_row, base_index)
    ours_key = _cell_key(ours_cell, ours_row, ours_index)
    theirs_key = _cell_key(theirs_cell, theirs_row, theirs_index)
    if ours_key == theirs_key:
        chosen = ours_cell if not ours_cell.is_empty() else theirs_cell
        if base_key != ours_key:
            result.changes.append(
                MergeChange(
                    sheet_name,
                    "cell_modified",
                    row=row_key,
                    column=column_name,
                    base=base_cell.display(include_formula=True),
                    ours=ours_cell.display(include_formula=True),
                    theirs=theirs_cell.display(include_formula=True),
                    result=chosen.display(include_formula=True),
                )
            )
        return chosen.clone()
    if ours_key == base_key:
        if theirs_key != base_key:
            result.changes.append(
                MergeChange(
                    sheet_name,
                    "cell_modified",
                    row=row_key,
                    column=column_name,
                    base=base_cell.display(include_formula=True),
                    ours=ours_cell.display(include_formula=True),
                    theirs=theirs_cell.display(include_formula=True),
                    result=theirs_cell.display(include_formula=True),
                    note="theirs",
                )
            )
        return theirs_cell.clone()
    if theirs_key == base_key:
        result.changes.append(
            MergeChange(
                sheet_name,
                "cell_modified",
                row=row_key,
                column=column_name,
                base=base_cell.display(include_formula=True),
                ours=ours_cell.display(include_formula=True),
                theirs=theirs_cell.display(include_formula=True),
                result=ours_cell.display(include_formula=True),
                note="ours",
            )
        )
        return ours_cell.clone()

    chosen_side = strategy_choice(result.strategy, "ours")
    chosen = ours_cell if chosen_side == "ours" else theirs_cell
    result.changes.append(
        MergeChange(
            sheet_name,
            "cell_conflict",
            row=row_key,
            column=column_name,
            base=base_cell.display(include_formula=True),
            ours=ours_cell.display(include_formula=True),
            theirs=theirs_cell.display(include_formula=True),
            result=chosen.display(include_formula=True),
            conflict=True,
            note=f"按 {result.strategy} 选择 {chosen_side}",
        )
    )
    return chosen.clone()


def _merge_rows(
    sheet_name: str,
    base: Sheet,
    ours: Sheet,
    theirs: Sheet,
    column_map: _ColumnMap,
    result: MergeResult,
) -> list[Row]:
    base_rows = _row_identity_map(base)
    ours_rows = _row_identity_map(ours)
    theirs_rows = _row_identity_map(theirs)
    identities = list(
        OrderedDict.fromkeys(
            [*ours_rows.keys(), *theirs_rows.keys(), *base_rows.keys()]
        ).keys()
    )
    output_by_identity: OrderedDict[tuple[str, int], tuple[Row, Mapping[int, int]]] = OrderedDict()

    reverse_ours = {value: key for key, value in column_map.ours.items()}
    reverse_theirs = {value: key for key, value in column_map.theirs.items()}
    reverse_base = {value: key for key, value in column_map.base.items()}

    for identity in identities:
        base_row = base_rows.get(identity)
        ours_row = ours_rows.get(identity)
        theirs_row = theirs_rows.get(identity)
        row_key = identity[0]
        if base_row is None:
            if ours_row is not None and theirs_row is None:
                output_by_identity[identity] = (
                    _remap_row(ours_row, column_map.ours, len(column_map.headers)),
                    column_map.ours,
                )
                result.changes.append(
                    MergeChange(sheet_name, "row_added", row=row_key, note="ours")
                )
                continue
            if ours_row is None and theirs_row is not None:
                output_by_identity[identity] = (
                    _remap_row(theirs_row, column_map.theirs, len(column_map.headers)),
                    column_map.theirs,
                )
                result.changes.append(
                    MergeChange(sheet_name, "row_added", row=row_key, note="theirs")
                )
                continue
            if ours_row is None and theirs_row is None:
                continue
        elif ours_row is None and theirs_row is None:
            result.changes.append(
                MergeChange(sheet_name, "row_deleted", row=row_key, note="both")
            )
            continue
        elif ours_row is None or theirs_row is None:
            present_side = "theirs" if ours_row is None else "ours"
            present_row = theirs_row if ours_row is None else ours_row
            changed = present_row is not None and _row_changed(
                base_row,
                present_row,
                base,
                theirs if present_side == "theirs" else ours,
                {
                    source: target
                    for target, source in (
                        column_map.theirs if present_side == "theirs" else column_map.ours
                    ).items()
                },
            )
            if not changed and present_row is not None:
                result.changes.append(
                    MergeChange(
                        sheet_name,
                        "row_deleted",
                        row=row_key,
                        note="ours" if ours_row is None else "theirs",
                    )
                )
                continue
            chosen_side = (
                "ours"
                if ours_row is not None
                else "theirs"
                if theirs_row is not None
                else strategy_choice(result.strategy, "ours")
            )
            chosen_row = ours_row if chosen_side == "ours" else theirs_row
            chosen_map = column_map.ours if chosen_side == "ours" else column_map.theirs
            if chosen_row is not None:
                output_by_identity[identity] = (
                    _remap_row(chosen_row, chosen_map, len(column_map.headers)),
                    chosen_map,
                )
            result.changes.append(
                MergeChange(
                    sheet_name,
                    "row_delete_conflict",
                    row=row_key,
                    conflict=True,
                    result=chosen_side,
                    note="一侧删除，另一侧修改",
                )
            )
            continue

        assert base_row is not None and ours_row is not None and theirs_row is not None
        merged_cells: list[Cell] = []
        for output_index, header in enumerate(column_map.headers):
            base_index = reverse_base.get(output_index)
            ours_index = reverse_ours.get(output_index)
            theirs_index = reverse_theirs.get(output_index)
            base_cell = (
                base_row.cells[base_index]
                if base_index is not None and base_index < len(base_row.cells)
                else Cell()
            )
            ours_cell = (
                ours_row.cells[ours_index]
                if ours_index is not None and ours_index < len(ours_row.cells)
                else Cell()
            )
            theirs_cell = (
                theirs_row.cells[theirs_index]
                if theirs_index is not None and theirs_index < len(theirs_row.cells)
                else Cell()
            )
            merged_cells.append(
                _merge_cell(
                    sheet_name,
                    row_key,
                    header.display_name(),
                    base_cell,
                    ours_cell,
                    theirs_cell,
                    base_row,
                    ours_row,
                    theirs_row,
                    base_index,
                    ours_index,
                    theirs_index,
                    result,
                )
            )
        output_by_identity[identity] = (
            Row(row_key, merged_cells, ours_row.source_row),
            {},
        )

    return _order_merged_rows(ours, theirs, output_by_identity)


def _order_merged_rows(
    ours: Sheet,
    theirs: Sheet,
    rows_by_identity: Mapping[tuple[str, int], tuple[Row, Mapping[int, int]]],
) -> list[Row]:
    """Use ours order as the base, then insert rows introduced by theirs."""
    identities: list[tuple[str, int]] = []
    for identity in ours.row_identities():
        if identity in rows_by_identity:
            identities.append(identity)
    seen = set(identities)

    theirs_identities = theirs.row_identities()
    for index, identity in enumerate(theirs_identities):
        if identity in seen or identity not in rows_by_identity:
            continue
        predecessor = next(
            (
                candidate
                for candidate in reversed(theirs_identities[:index])
                if candidate in seen
            ),
            None,
        )
        if predecessor is not None:
            identities.insert(identities.index(predecessor) + 1, identity)
        else:
            successor = next(
                (
                    candidate
                    for candidate in theirs_identities[index + 1 :]
                    if candidate in seen
                ),
                None,
            )
            if successor is not None:
                identities.insert(identities.index(successor), identity)
            else:
                identities.append(identity)
        seen.add(identity)

    for identity in rows_by_identity:
        if identity not in seen:
            identities.append(identity)
            seen.add(identity)
    return [rows_by_identity[identity][0] for identity in identities]


def format_merge_text(result: MergeResult) -> str:
    lines = [
        f"merge strategy: {result.strategy}",
        f"changes: {len(result.changes)}",
        f"conflicts: {len(result.conflicts)}",
    ]
    for change in result.changes:
        marker = "!" if change.conflict else "~"
        location = ".".join(
            value for value in (change.row, change.column) if value
        )
        parts = [
            f"{marker} {change.kind}",
            f"{change.sheet}!{location}" if location else change.sheet,
        ]
        if change.base or change.ours or change.theirs:
            parts.append(
                f"base={change.base!r} ours={change.ours!r} theirs={change.theirs!r}"
            )
        parts.append(f"result={change.result!r}")
        if change.note:
            parts.append(f"note={change.note}")
        lines.append(" | ".join(parts))
    if result.issues:
        lines.append("")
        lines.append("issues:")
        for issue in result.issues:
            lines.append(f"- [{issue.severity}] {issue.code}: {issue.message}")
    return "\n".join(lines)


def write_merge_excel(
    result: MergeResult,
    output_path: str | Path,
    *,
    allow_conflicts: bool = False,
) -> str:
    if result.conflicts and not allow_conflicts:
        result.require_resolved()
    output_path = Path(output_path)
    result.document.save(output_path)

    wb = load_workbook(output_path)
    review = wb.create_sheet("Review", 0)
    conflicts_sheet = wb.create_sheet("Conflicts", 1)
    issues_sheet = wb.create_sheet("Issues", 2)
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(color="FFFFFF", bold=True)
    conflict_fill = PatternFill("solid", fgColor="FFC7CE")
    change_fill = PatternFill("solid", fgColor="FFF2CC")

    review.append(["Metric", "Value"])
    review.append(["Strategy", result.strategy])
    review.append(["Changes", len(result.changes)])
    review.append(["Conflicts", len(result.conflicts)])
    review.append(["OK", result.ok])
    for cell in review[1]:
        cell.fill = header_fill
        cell.font = header_font
    review.column_dimensions["A"].width = 20
    review.column_dimensions["B"].width = 80

    conflicts_sheet.append(
        ["Sheet", "Row", "Column", "Base", "Ours", "Theirs", "Result", "Note"]
    )
    for change in result.conflicts:
        conflicts_sheet.append(
            [
                change.sheet,
                change.row,
                change.column,
                change.base,
                change.ours,
                change.theirs,
                change.result,
                change.note,
            ]
        )
    for row_cells in conflicts_sheet.iter_rows(min_row=2):
        for cell in row_cells:
            cell.fill = conflict_fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in conflicts_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    conflicts_sheet.freeze_panes = "A2"
    conflicts_sheet.auto_filter.ref = conflicts_sheet.dimensions

    changes_sheet = wb.create_sheet("Changes", 3)
    changes_sheet.append(
        ["Sheet", "Kind", "Row", "Column", "Base", "Ours", "Theirs", "Result", "Conflict", "Note"]
    )
    for change in result.changes:
        changes_sheet.append(
            [
                change.sheet,
                change.kind,
                change.row,
                change.column,
                change.base,
                change.ours,
                change.theirs,
                change.result,
                change.conflict,
                change.note,
            ]
        )
    for row_cells in changes_sheet.iter_rows(min_row=2):
        fill = conflict_fill if row_cells[8].value else change_fill
        for cell in row_cells:
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in changes_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    changes_sheet.freeze_panes = "A2"
    changes_sheet.auto_filter.ref = changes_sheet.dimensions

    issues_sheet.append(["Severity", "Code", "Sheet", "Row", "Column", "Message", "Reference"])
    for issue in result.issues:
        issues_sheet.append(
            [
                issue.severity,
                issue.code,
                issue.sheet or "",
                issue.row or "",
                issue.column or "",
                issue.message,
                issue.reference or "",
            ]
        )
    for cell in issues_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    issues_sheet.freeze_panes = "A2"
    issues_sheet.auto_filter.ref = issues_sheet.dimensions

    widths = [18, 24, 18, 26, 36, 42, 42, 42, 12, 44]
    for sheet in (conflicts_sheet, changes_sheet):
        for index, width in enumerate(widths[: len(sheet[1])], start=1):
            sheet.column_dimensions[chr(64 + index)].width = width
    wb.save(output_path)
    wb.close()
    return str(output_path)
