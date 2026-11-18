use std::collections::HashSet;
use std::sync::OnceLock;

use regex::Regex;
use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
pub struct FormulaFinding {
    pub code: String,
    pub severity: String,
    pub message: String,
    pub reference: Option<String>,
}

fn external_re() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r"\[[^\]]+\]").expect("external ref regex"))
}

fn function_re() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r"(?i)([A-Z][A-Z0-9_.]*)\s*\(").expect("function regex"))
}

fn sheet_ref_re() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(
            r#"(?:"'([^']|'')+'|([A-Za-z0-9_#\p{Han}]+))!\$?([A-Z]{1,3})(\$?)(\d+)(?::\$?([A-Z]{1,3})(\$?)(\d+))?"#,
        )
        .expect("sheet reference regex")
    })
}

fn bare_ref_re() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(r"\$?([A-Z]{1,3})(\$?)(\d+)(?::\$?([A-Z]{1,3})(\$?)(\d+))?")
            .expect("cell reference regex")
    })
}

fn issue(code: &str, severity: &str, message: &str, reference: Option<&str>) -> FormulaFinding {
    FormulaFinding {
        code: code.to_string(),
        severity: severity.to_string(),
        message: message.to_string(),
        reference: reference.map(ToOwned::to_owned),
    }
}

pub fn analyze_formula(formula: &str, source_row: usize) -> Vec<FormulaFinding> {
    let mut findings = Vec::new();
    let normalized = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    let formula = normalized.as_str();

    if external_re().is_match(formula) {
        findings.push(issue(
            "external_workbook_ref",
            "error",
            "检测到跨工作簿引用；公式缓存、外链和保存兼容性风险较高。",
            Some(formula),
        ));
        return findings;
    }

    const COMPLEX_FUNCTIONS: &[&str] = &[
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
    ];
    let mut complex = HashSet::new();
    for captures in function_re().captures_iter(formula) {
        let name = captures[1].to_ascii_uppercase();
        if COMPLEX_FUNCTIONS.contains(&name.as_str()) {
            complex.insert(name);
        }
    }
    if !complex.is_empty() {
        let mut names: Vec<_> = complex.into_iter().collect();
        names.sort();
        findings.push(issue(
            "complex_function",
            "warning",
            &format!(
                "检测到动态/易失或计算上下文敏感的函数: {}。",
                names.join(", ")
            ),
            None,
        ));
    }

    let mut sheet_spans = Vec::new();
    for captures in sheet_ref_re().captures_iter(formula) {
        let whole = captures.get(0).expect("whole reference").as_str();
        sheet_spans.push((
            captures.get(0).expect("whole reference").start(),
            captures.get(0).expect("whole reference").end(),
        ));
        let after_bang = whole
            .find('!')
            .map(|index| &whole[index + 1..])
            .unwrap_or_default();
        let first_abs_col = after_bang.starts_with('$');
        let end_abs_col = captures.get(7).is_some_and(|value| value.as_str() == "$");
        let first_abs_row = captures.get(4).is_some_and(|value| value.as_str() == "$");
        let end_abs_row = captures.get(8).is_some_and(|value| value.as_str() == "$");
        let fixed = first_abs_col
            && first_abs_row
            && (captures.get(7).is_none() || (end_abs_col && end_abs_row));
        if !fixed {
            findings.push(issue(
                "cross_sheet_relative_ref",
                "error",
                "跨页引用必须是固定目标（行、列均使用 $）。",
                Some(whole),
            ));
        }
    }

    for captures in bare_ref_re().captures_iter(formula) {
        let start = captures.get(0).unwrap().start();
        if start > 0 {
            let previous = formula[..start].chars().next_back().unwrap_or(' ');
            if previous.is_alphanumeric() || previous == '_' || previous == '$' {
                continue;
            }
        }
        let prefix = &formula[..start];
        let segment_start = prefix
            .rfind(|character| {
                matches!(
                    character,
                    ',' | '(' | '+' | '-' | '*' | '/' | '&' | '=' | '<' | '>'
                )
            })
            .map(|index| index + 1)
            .unwrap_or(0);
        if prefix[segment_start..].contains('!') {
            continue;
        }
        if sheet_spans
            .iter()
            .any(|(span_start, span_end)| start >= *span_start && start < *span_end)
        {
            continue;
        }
        let after = formula[captures.get(0).unwrap().end()..].trim_start();
        if after.starts_with('(') {
            continue;
        }
        let first_row: usize = captures[3].parse().unwrap_or(0);
        let end_row = captures
            .get(6)
            .and_then(|value| value.as_str().parse().ok())
            .unwrap_or(first_row);
        if first_row != source_row || end_row != source_row {
            findings.push(issue(
                "cross_row_ref",
                "error",
                "公式引用了其他数据行；只允许行内引用或跨页固定目标。",
                captures.get(0).map(|value| value.as_str()),
            ));
        }
    }

    let mut seen = HashSet::new();
    findings.retain(|finding| seen.insert((finding.code.clone(), finding.reference.clone())));
    findings
}

fn column_number(column: &str) -> usize {
    column.bytes().fold(0usize, |value, byte| {
        value * 26 + (byte - b'A' + 1) as usize
    })
}

fn column_name(mut column: usize) -> String {
    let mut result = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        result.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    result
}

fn inside_string(formula: &str, position: usize) -> bool {
    let mut quoted = false;
    let mut index = 0usize;
    for character in formula.chars() {
        let char_len = character.len_utf8();
        if index >= position {
            break;
        }
        if character == '"' {
            quoted = !quoted;
        }
        index += char_len;
    }
    quoted
}

fn is_bare_cell_reference(formula: &str, start: usize, end: usize) -> bool {
    if inside_string(formula, start) {
        return false;
    }
    if start > 0 {
        let previous = formula[..start].chars().next_back().unwrap_or(' ');
        if previous.is_alphanumeric() || previous == '_' || previous == '$' {
            return false;
        }
    }
    let prefix = &formula[..start];
    let segment_start = prefix
        .rfind(|character| {
            matches!(
                character,
                ',' | '(' | '+' | '-' | '*' | '/' | '&' | '=' | '<' | '>'
            )
        })
        .map(|index| index + 1)
        .unwrap_or(0);
    if prefix[segment_start..].contains('!') {
        return false;
    }
    !formula[end..].trim_start().starts_with('(')
}

/// Rebind row-local relative references after a row/column move.
///
/// Cross-sheet and external references are left untouched. Absolute row
/// components are preserved. Column mapping is applied when provided.
pub fn bind_formula(
    formula: &str,
    source_row: usize,
    target_row: usize,
    column_map: Option<&[Option<usize>]>,
) -> String {
    let formula = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    let row_delta = target_row as isize - source_row as isize;
    let mut replacements = Vec::new();
    for captures in bare_ref_re().captures_iter(&formula) {
        let whole = captures.get(0).expect("whole reference");
        if !is_bare_cell_reference(&formula, whole.start(), whole.end()) {
            continue;
        }
        let start_abs_col = whole.as_str().starts_with('$');
        let start_row_abs = captures.get(2).is_some_and(|value| value.as_str() == "$");
        let start_col = column_number(&captures[1]);
        let start_row: usize = captures[3].parse().unwrap_or(0);
        let start_output_col = column_map
            .and_then(|map| map.get(start_col.saturating_sub(1)).copied().flatten())
            .map(|index| index + 1)
            .unwrap_or(start_col);
        let start_output_row = if start_row_abs {
            start_row
        } else {
            (start_row as isize + row_delta).max(1) as usize
        };
        let mut replacement = String::new();
        if start_abs_col {
            replacement.push('$');
        }
        replacement.push_str(&column_name(start_output_col));
        if start_row_abs {
            replacement.push('$');
        }
        replacement.push_str(&start_output_row.to_string());

        if let Some(end_col) = captures.get(4) {
            let end_col_text = whole
                .as_str()
                .split_once(':')
                .map(|(_, value)| value)
                .unwrap_or(end_col.as_str());
            let end_abs_col = end_col_text.starts_with('$');
            let end_row_abs = captures.get(5).is_some_and(|value| value.as_str() == "$");
            let end_col = column_number(end_col.as_str());
            let end_row: usize = captures[6].parse().unwrap_or(0);
            let end_output_col = column_map
                .and_then(|map| map.get(end_col.saturating_sub(1)).copied().flatten())
                .map(|index| index + 1)
                .unwrap_or(end_col);
            let end_output_row = if end_row_abs {
                end_row
            } else {
                (end_row as isize + row_delta).max(1) as usize
            };
            replacement.push(':');
            if end_abs_col {
                replacement.push('$');
            }
            replacement.push_str(&column_name(end_output_col));
            if end_row_abs {
                replacement.push('$');
            }
            replacement.push_str(&end_output_row.to_string());
        }
        if replacement != whole.as_str() {
            replacements.push((whole.start(), whole.end(), replacement));
        }
    }
    let mut result = formula;
    for (start, end, replacement) in replacements.into_iter().rev() {
        result.replace_range(start..end, &replacement);
    }
    result
}

pub fn formula_signature(formula: &str, source_row: usize) -> String {
    bind_formula(formula, source_row, 4, None)
}
