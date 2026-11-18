"""Formula safety checks and safe rebinding for simple configuration formulas.

Supported safe forms:

* row-local references, such as ``=C4&D4`` in row 4
* row-local ranges, such as ``=TEXTJOIN(",",TRUE,C4:D4)`` in row 4
* fixed cross-sheet references, such as ``='#Attr'!$A$1`` or ``=Sheet1!$B$2``

Anything more complex is reported as a risk. Nothing in this module claims to
implement Excel's complete formula grammar.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from formula_utils import normalize_formula_for_compare, parse_cell_refs


_EXTERNAL_REF_RE = re.compile(r"\[[^\]]+\.")
_DYNAMIC_OR_VOLATILE_FUNCTIONS = {
    "BYCOL",
    "BYROW",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "FILTER",
    "HSTACK",
    "INDIRECT",
    "LAMBDA",
    "LET",
    "MAP",
    "OFFSET",
    "REDUCE",
    "SCAN",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "TAKE",
    "TOCOL",
    "TOROW",
    "UNIQUE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
}


@dataclass(frozen=True, slots=True)
class FormulaFinding:
    code: str
    severity: str
    message: str
    reference: str | None = None

    @property
    def is_error(self) -> bool:
        return self.severity == "error"


def formula_signature(formula: str | None, source_row: int, source_col: int) -> str:
    """Return a location-independent signature for a simple formula."""
    if not formula:
        return ""
    return normalize_formula_for_compare(formula, source_row, source_col)


def _function_names(formula: str) -> set[str]:
    return {match.upper() for match in re.findall(r"(?<![A-Z0-9_.])([A-Z][A-Z0-9_.]*)\s*\(", formula, re.I)}


def analyze_formula(
    formula: str | None,
    *,
    source_row: int,
    source_col: int,
    is_array: bool = False,
) -> list[FormulaFinding]:
    """Classify a formula against the project's simple-formula contract."""
    if not formula:
        return []
    findings: list[FormulaFinding] = []
    if not str(formula).startswith("="):
        findings.append(
            FormulaFinding(
                "formula_text_invalid",
                "warning",
                "单元格看起来是公式，但没有以 '=' 开头。",
                str(formula)[:80],
            )
        )
        return findings

    if is_array:
        findings.append(
            FormulaFinding(
                "array_formula",
                "warning",
                "检测到数组公式；重构后的合并/写入流程不保证其动态数组元数据完整。",
            )
        )

    if _EXTERNAL_REF_RE.search(formula):
        findings.append(
            FormulaFinding(
                "external_workbook_ref",
                "error",
                "检测到跨工作簿引用；公式缓存、外链和保存兼容性风险较高。",
                formula[:120],
            )
        )

    functions = _function_names(formula) & _DYNAMIC_OR_VOLATILE_FUNCTIONS
    if functions:
        findings.append(
            FormulaFinding(
                "complex_function",
                "warning",
                "检测到动态/易失或计算上下文敏感的函数: "
                + ", ".join(sorted(functions))
                + "。",
            )
        )

    for ref in parse_cell_refs(formula):
        display_ref = ref.get("ref") or ""
        if ref.get("file"):
            # Already covered by the external-workbook finding.
            continue
        if ref.get("sheet"):
            fixed_target = bool(ref.get("row_start_abs") and ref.get("col_start_abs"))
            if ref.get("is_range"):
                fixed_target = fixed_target and bool(
                    ref.get("row_end_abs") and ref.get("col_end_abs")
                )
            if not fixed_target:
                findings.append(
                    FormulaFinding(
                        "cross_sheet_relative_ref",
                        "error",
                        "跨页引用必须是固定目标（行、列均使用 $）。",
                        display_ref,
                    )
                )
            continue

        rows = {int(ref["row_start"])}
        if ref.get("is_range"):
            rows.add(int(ref["row_end"]))
        if rows != {int(source_row)}:
            findings.append(
                FormulaFinding(
                    "cross_row_ref",
                    "error",
                    "公式引用了其他数据行；只允许行内引用或跨页固定目标。",
                    display_ref,
                )
            )

    # De-duplicate findings caused by multiple references.
    unique: list[FormulaFinding] = []
    seen: set[tuple[str, str | None]] = set()
    for finding in findings:
        key = (finding.code, finding.reference)
        if key not in seen:
            unique.append(finding)
            seen.add(key)
    return unique


def _rebuild_cell_ref(ref: dict, row: int, col: int) -> str:
    col_text = ("$" if ref.get("col_start_abs") else "") + _number_to_col(col)
    row_text = ("$" if ref.get("row_start_abs") else "") + str(row)
    return col_text + row_text


def _rebuild_range_end(ref: dict, row: int, col: int) -> str:
    col_text = ("$" if ref.get("col_end_abs") else "") + _number_to_col(col)
    row_text = ("$" if ref.get("row_end_abs") else "") + str(row)
    return col_text + row_text


def _number_to_col(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _col_to_number(column: str) -> int:
    result = 0
    for char in column.replace("$", "").upper():
        result = result * 26 + (ord(char) - 64)
    return result


def bind_row_local_formula(
    formula: str | None,
    *,
    source_row: int,
    source_col: int,
    target_row: int,
    target_col: int,
) -> str | None:
    """Shift bare/relative references when a cell moves to another coordinate.

    Cross-sheet and external references are left untouched. Absolute local
    components are also left untouched.
    """
    if not formula or source_row <= 0 or source_col <= 0:
        return formula
    if target_row == source_row and target_col == source_col:
        return formula

    result = formula
    refs = parse_cell_refs(formula)
    for ref in sorted(refs, key=lambda item: item["position"], reverse=True):
        if ref.get("sheet") or ref.get("file"):
            continue

        start_row = int(ref["row_start"])
        start_col = _col_to_number(str(ref["col_start"]))
        if not ref.get("row_start_abs"):
            start_row = target_row + (start_row - source_row)
        if not ref.get("col_start_abs"):
            start_col = target_col + (start_col - source_col)
        start_text = _rebuild_cell_ref(ref, start_row, start_col)

        if ref.get("is_range"):
            end_row = int(ref["row_end"])
            end_col = _col_to_number(str(ref["col_end"]))
            if not ref.get("row_end_abs"):
                end_row = target_row + (end_row - source_row)
            if not ref.get("col_end_abs"):
                end_col = target_col + (end_col - source_col)
            replacement = start_text + ":" + _rebuild_range_end(ref, end_row, end_col)
        else:
            replacement = start_text

        position = int(ref["position"])
        original = ref["ref"]
        result = result[:position] + replacement + result[position + len(original):]
    return result


def formula_references_sheet_column(
    formula: str | None,
    *,
    sheet_name: str,
    column_index: int,
) -> bool:
    """Whether a formula references the given same-sheet column."""
    if not formula:
        return False
    for ref in parse_cell_refs(formula):
        if ref.get("file"):
            continue
        if ref.get("sheet") and ref["sheet"] != sheet_name:
            continue
        cols = {_col_to_number(str(ref["col_start"]))}
        if ref.get("is_range"):
            cols.add(_col_to_number(str(ref["col_end"])))
        if column_index in cols:
            return True
    return False
