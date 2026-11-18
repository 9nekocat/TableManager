# -*- coding: utf-8 -*-
"""
Formula utilities for DataTable merge.
- Parse cell/range references from formulas
- Display formula with value: "=FORMULA -> VALUE"
- Adjust references after row insertion/deletion during merge
"""

import re
import string


def _col_to_num(col_str):
    """Convert column letters to 1-based number. A=1, B=2, ..., Z=26, AA=27."""
    n = 0
    for c in col_str.upper():
        if "A" <= c <= "Z":
            n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def _num_to_col(n):
    """Convert 1-based column number to letters. 1=A, 27=AA."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(rem + ord("A")) + result
    return result


# Match cell or range: $?A-Z letters $?digits  with optional :$?A-Z letters $?digits
_CELL_RE = re.compile(r"(\$?[A-Za-z]{1,3})(\$?\d+)")
_RANGE_RE = re.compile(r"(\$?[A-Za-z]{1,3})(\$?\d+):(\$?[A-Za-z]{1,3})(\$?\d+)")

# Matches [file]'Sheet'! or [file]Sheet! or 'Sheet'! or Sheet!
_SHEET_BANG_RE = re.compile(
    r"(?:\[([^\]]+)\])?"      # optional [file.xlsx]
    r"(?:'([^']+)')"          # 'quoted sheet name'
    r"!"                       # bang
)
_SHEET_BANG_SIMPLE_RE = re.compile(
    r"(?:\[([^\]]+)\])?"      # optional [file.xlsx]
    # simple sheet name; also allows CJK etc. (first char not a digit so a
    # bare number like 1!A1 is not treated as a sheet name)
    r"([^'! ,\(\)=+\-*/^&<>\"$\d][^'! ,\(\)=+\-*/^&<>\"$]*)"
    r"!"                       # bang
)


def _parse_cell(ref_str):
    """Parse a single cell reference like A1 or $A$1."""
    m = re.match(r"^(\$?)([A-Za-z]+)(\$?)(\d+)$", ref_str)
    if not m:
        return None
    return {
        "col_abs": m.group(1) == "$",
        "col": m.group(2).upper(),
        "row_abs": m.group(3) == "$",
        "row": int(m.group(4)),
    }


def _string_spans(formula):
    """Return (start, end) spans of "..." string literals in a formula.

    Inside an Excel string, "" is an escaped quote. Cell-like text inside
    string literals (e.g. ="B4"&B4) must not be treated as references.
    """
    spans = []
    i = 0
    n = len(formula)
    while i < n:
        if formula[i] == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            spans.append((i, min(j + 1, n)))
            i = j + 1
        else:
            i += 1
    return spans


def parse_cell_refs(formula):
    """
    Parse a formula string and extract all cell/range references.

    Strategy: use a stateful scanner.
    - When we see '!' we capture the preceding sheet/file prefix.
    - All cell/range refs after '!' inherit that prefix until next operator.
    - Cell refs not after a '!' are bare (no sheet context).

    Returns list of dicts with keys: ref, sheet, file, col_start, row_start,
    col_end, row_end, col_abs, row_abs, col_end_abs, row_end_abs, position
    """
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return []

    # String literal spans: anything inside "..." is not a real reference
    string_spans = _string_spans(formula)

    def _in_string(pos):
        return any(s <= pos < e for s, e in string_spans)

    # Step 1: find all sheet! markers and their positions
    # marker: (prefix_start, bang_pos, ref_end, file, sheet, quoted)
    markers = []
    covered_ranges = []  # (start, end) of sheet_name!cell_area segments

    # Quoted sheet names
    for m in _SHEET_BANG_RE.finditer(formula):
        if _in_string(m.start()):
            continue
        file_name, sheet_name = m.groups()
        bang_pos = m.end()
        # Find how far the cell ref extends
        end = bang_pos
        while end < len(formula) and formula[end] in "$ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789:":
            end += 1
        covered_ranges.append((m.start(), end))
        markers.append((m.start(), bang_pos, end, file_name, sheet_name, True))

    # Simple sheet names (skip those already in covered_ranges)
    for m in _SHEET_BANG_SIMPLE_RE.finditer(formula):
        if _in_string(m.start()):
            continue
        if any(r[0] <= m.start() < r[1] for r in covered_ranges):
            continue
        file_name, sheet_name = m.groups()
        bang_pos = m.end()
        end = bang_pos
        while end < len(formula) and formula[end] in "$ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789:":
            end += 1
        covered_ranges.append((m.start(), end))
        markers.append((m.start(), bang_pos, end, file_name, sheet_name, False))

    # Step 2: collect all cell/range refs
    refs = []

    # Range references
    for m in _RANGE_RE.finditer(formula):
        pos = m.start()
        if _in_string(pos):
            continue
        # Skip if inside a sheet name part (before !)
        skip = False
        for (cov_start, cov_end), marker in zip(covered_ranges, markers):
            if cov_start <= pos < marker[1]:  # marker[1] = bang_pos
                skip = True
                break
        if skip:
            continue

        col1, row1, col2, row2 = m.group(1), m.group(2), m.group(3), m.group(4)
        a = _parse_cell(col1 + row1)
        b = _parse_cell(col2 + row2)
        if a is None or b is None:
            continue

        full_ref = f"{col1}{row1}:{col2}{row2}"
        file_name, sheet_name, prefix_start, quoted = _find_marker(markers, pos)
        # If this ref has a sheet prefix, position is the prefix start recorded
        # at regex match time (no error-prone backtracking).
        actual_pos = pos
        if sheet_name and prefix_start is not None:
            actual_pos = prefix_start
        refs.append(_make_ref(full_ref, a, b, file_name, sheet_name, actual_pos, quoted))

    # Single cell references (skip those already part of a range)
    range_positions = set()
    for m in _RANGE_RE.finditer(formula):
        range_positions.add(m.start())
        # Also mark the second cell pos (after ':')
        rpos = m.start() + m.group(1).__len__() + m.group(2).__len__() + 1  # +1 for ':'
        range_positions.add(rpos)

    for m in _CELL_RE.finditer(formula):
        pos = m.start()
        if pos in range_positions:
            continue
        if _in_string(pos):
            continue

        # Skip if inside a sheet name part (before !)
        skip = False
        for (cov_start, cov_end), marker in zip(covered_ranges, markers):
            if cov_start <= pos < marker[1]:  # marker[1] = bang_pos
                skip = True
                break
        if skip:
            continue

        # Skip function names that look like cell refs (LOG10(, ATAN2(), ...):
        # if the next non-'$' char after the match is '(', it is a function call.
        fend = m.end()
        while fend < len(formula) and formula[fend] == "$":
            fend += 1
        if fend < len(formula) and formula[fend] == "(":
            continue

        col_str, row_str = m.group(1), m.group(2)
        a = _parse_cell(col_str + row_str)
        if a is None:
            continue

        full_ref = f"{col_str}{row_str}"
        file_name, sheet_name, prefix_start, quoted = _find_marker(markers, pos)
        actual_pos = pos
        if sheet_name and prefix_start is not None:
            actual_pos = prefix_start
        refs.append(_make_ref(full_ref, a, a, file_name, sheet_name, actual_pos, quoted))

    return refs


def _find_marker(markers, pos):
    """Find which sheet marker covers this position (after !).
    Returns (file_name, sheet_name, prefix_start, quoted)."""
    for (prefix_start, bang_pos, end, file_name, sheet_name, quoted) in markers:
        if bang_pos <= pos < end:
            return file_name, sheet_name, prefix_start, quoted
    return None, None, None, False


def _sheet_prefix(file_name, sheet_name, quoted):
    """Rebuild the [file]sheet! prefix, restoring quotes when the sheet name
    was quoted in the original text ('' escapes an inner quote)."""
    prefix = ""
    if file_name:
        prefix += f"[{file_name}]"
    if sheet_name:
        if quoted:
            prefix += "'" + sheet_name.replace("'", "''") + "'!"
        else:
            prefix += f"{sheet_name}!"
    return prefix


def _make_ref(full_ref, a, b, file_name, sheet_name, pos, quoted=False):
    a_ref = (f"${a['col']}" if a["col_abs"] else a["col"]) + \
            (f"${a['row']}" if a["row_abs"] else str(a["row"]))
    b_ref = (f"${b['col']}" if b["col_abs"] else b["col"]) + \
            (f"${b['row']}" if b["row_abs"] else str(b["row"]))
    is_range = a_ref != b_ref

    # Build the complete matched ref including sheet/file prefix, preserving
    # the original quoting so len(ref) matches the actual source text.
    complete_ref = _sheet_prefix(file_name, sheet_name, quoted) + full_ref

    return {
        "ref": complete_ref,         # full matched text (e.g., 'Sheet 2'!B4:D20)
        "cell_ref": full_ref,        # just the cell part (e.g., B4:D20)
        "sheet": sheet_name,
        "file": file_name,
        "quoted": quoted,
        "col_start": a["col"],
        "col_start_abs": a["col_abs"],
        "row_start": a["row"],
        "row_start_abs": a["row_abs"],
        "col_end": b["col"],
        "col_end_abs": b["col_abs"],
        "row_end": b["row"],
        "row_end_abs": b["row_abs"],
        "is_range": is_range,
        "position": pos,
    }


def format_display(formula, value):
    """
    Format a cell for merge preview display.
    Prefixes with ' (apostrophe) so Excel treats it as literal text,
    not as a formula that would show #NAME? errors.
    If formula exists: "'=SUM(A1:A10) -> 42"
    If only value: just the value (prefixed with ' if it starts with =)
    """
    fml = str(formula).strip() if formula is not None else ""
    val = str(value).strip() if value is not None else ""

    if fml.startswith("="):
        if val:
            return f"'{fml} → {val}"
        return f"'{fml}"

    # Values starting with = would also confuse Excel
    if val.startswith("="):
        return f"'{val}"
    return val


def adjust_formula(formula, row_map, col_map=None):
    """
    Adjust cell/range references in a formula based on row mapping.

    row_map: dict {old_row: new_row}  or  list of (old_row, new_row) pairs
             Rows not in the map are preserved as-is.
             If old_row -> None or -1, the row was deleted.

    col_map: dict {old_col_letter: new_col_letter}, optional.

    Returns adjusted formula string.
    """
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return formula

    if isinstance(row_map, list):
        row_map = dict(row_map)

    refs = parse_cell_refs(formula)
    if not refs:
        return formula

    # Build replacement segments, working from right to left to preserve positions
    result = formula
    for ref_info in sorted(refs, key=lambda r: r["position"], reverse=True):
        old_ref = ref_info["ref"]
        new_ref = _adjust_single_ref(ref_info, row_map, col_map)
        if new_ref != old_ref:
            pos = ref_info["position"]
            result = result[:pos] + new_ref + result[pos + len(old_ref):]

    return result


def normalize_formula_for_compare(formula, source_row, source_col):
    """Return a location-independent signature for local relative references.

    Excel updates relative A1 references when rows are inserted.  For example,
    ``=B42`` in ``A42`` and ``=B43`` in ``A43`` are the same formula copied to
    a different row.  Comparing their raw strings creates noisy merge changes
    when a workbook has no cached formula values.

    Only *bare* (same-sheet) references are converted to R1C1-style tokens.
    References with an explicit sheet or workbook remain untouched: their
    target can legitimately change independently of this cell's position.
    Absolute row/column components are retained as absolute values, so a real
    change such as ``$B$4`` -> ``$B$5`` is still detected.
    """
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return formula
    if not isinstance(source_row, int) or not isinstance(source_col, int):
        return formula

    def r1c1_part(col, col_abs, row, row_abs):
        col_num = _col_to_num(col)
        row_part = f"R{row}" if row_abs else (
            "R" if row == source_row else f"R[{row - source_row}]"
        )
        col_part = f"C{col_num}" if col_abs else (
            "C" if col_num == source_col else f"C[{col_num - source_col}]"
        )
        return row_part + col_part

    result = formula
    refs = parse_cell_refs(formula)
    for ref_info in sorted(refs, key=lambda item: item["position"], reverse=True):
        # Keeping explicit sheet/file prefixes intact avoids turning a changed
        # cross-table reference into a false "same" result.
        if ref_info["sheet"] or ref_info["file"]:
            continue

        replacement = r1c1_part(
            ref_info["col_start"], ref_info["col_start_abs"],
            ref_info["row_start"], ref_info["row_start_abs"],
        )
        if ref_info["is_range"]:
            replacement += ":" + r1c1_part(
                ref_info["col_end"], ref_info["col_end_abs"],
                ref_info["row_end"], ref_info["row_end_abs"],
            )

        pos = ref_info["position"]
        original = ref_info["ref"]
        result = result[:pos] + replacement + result[pos + len(original):]

    return result


def _adjust_single_ref(ref_info, row_map, col_map):
    """Adjust a single cell or range reference. Returns the new full ref string."""

    # Cross-sheet / cross-file references must not be adjusted with this
    # sheet's row_map; return the original text (quotes preserved) unchanged.
    if ref_info["sheet"] or ref_info["file"]:
        return ref_info["ref"]

    def _adjust_row(row, is_abs):
        if is_abs:
            return row
        if row not in row_map:
            return row  # untouched rows: keep original
        new = row_map[row]
        if new is None or new < 0:
            return None  # row was deleted -> #REF!
        return new

    def _adjust_col(col, is_abs):
        if is_abs or col_map is None:
            return col
        new = col_map.get(col)
        return new if new else col

    def _rebuild(col, col_abs, row, row_abs):
        if row is None:
            return "#REF!"  # referenced row was deleted
        c = f"${col}" if col_abs else col
        r = f"${row}" if row_abs else row
        return f"{c}{r}"

    new_start = _rebuild(
        _adjust_col(ref_info["col_start"], ref_info["col_start_abs"]),
        ref_info["col_start_abs"],
        _adjust_row(ref_info["row_start"], ref_info["row_start_abs"]),
        ref_info["row_start_abs"],
    )

    if ref_info["is_range"]:
        new_end = _rebuild(
            _adjust_col(ref_info["col_end"], ref_info["col_end_abs"]),
            ref_info["col_end_abs"],
            _adjust_row(ref_info["row_end"], ref_info["row_end_abs"]),
            ref_info["row_end_abs"],
        )
        cell_part = f"{new_start}:{new_end}"
    else:
        cell_part = new_start

    # Rebuild the full ref with sheet/file prefix (quotes preserved). Note:
    # sheet/file refs already returned unchanged above, so this is normally
    # the bare-ref path with an empty prefix.
    prefix = _sheet_prefix(ref_info["file"], ref_info["sheet"],
                           ref_info.get("quoted", False))

    return prefix + cell_part


def build_row_map(old_rows, new_rows):
    """
    Build row mapping from old row indices to new row indices.
    Used when rows are added/removed during merge.

    old_rows: sorted list of row numbers in old version
    new_rows: sorted list of row numbers in new version
    Returns: dict {old_row: new_row}; rows deleted in the new version
             (trailing old rows beyond the new row count) map to None
    """
    mapping = {}
    old_set = set(old_rows)
    new_set = set(new_rows)

    # Simple case: rows are by position
    # For data rows (row >= 4), map by position order
    old_data = sorted([r for r in old_rows if r >= 4])
    new_data = sorted([r for r in new_rows if r >= 4])

    # If counts differ, map first N rows where N = min(len(old), len(new))
    for i in range(min(len(old_data), len(new_data))):
        mapping[old_data[i]] = new_data[i]

    # Trailing old rows with no counterpart were deleted -> map to None so
    # adjust_formula rewrites references to them as #REF!
    for i in range(len(new_data), len(old_data)):
        mapping[old_data[i]] = None

    return mapping
