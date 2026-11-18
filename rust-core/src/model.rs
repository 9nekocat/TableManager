use std::collections::HashMap;
use std::path::Path;

use calamine::{Data, Reader, Sheets, open_workbook_auto};
use serde::Serialize;

use crate::error::{CoreError, Result};
use crate::formula::analyze_formula;

pub const DATA_START_ROW: usize = 4;
const PLACEHOLDER_HEADERS: &[&str] = &["", "none", "null", "nan", "-"];

#[derive(Debug, Clone, Serialize)]
pub struct Issue {
    pub code: String,
    pub severity: String,
    pub message: String,
    pub sheet: Option<String>,
    pub row: Option<usize>,
    pub column: Option<String>,
    pub reference: Option<String>,
}

impl Issue {
    pub fn new(
        code: &str,
        severity: &str,
        message: impl Into<String>,
        sheet: Option<&str>,
        row: Option<usize>,
        column: Option<&str>,
        reference: Option<String>,
    ) -> Self {
        Self {
            code: code.to_string(),
            severity: severity.to_string(),
            message: message.into(),
            sheet: sheet.map(ToOwned::to_owned),
            row,
            column: column.map(ToOwned::to_owned),
            reference,
        }
    }

    pub fn is_error(&self) -> bool {
        self.severity == "error"
    }
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct Header {
    pub cn: String,
    pub en: String,
    #[serde(rename = "type")]
    pub type_name: String,
}

impl Header {
    pub fn new(cn: impl Into<String>, en: impl Into<String>, type_name: impl Into<String>) -> Self {
        Self {
            cn: cn.into().trim().to_string(),
            en: en.into().trim().to_string(),
            type_name: type_name.into().trim().to_string(),
        }
    }

    pub fn identity(&self) -> String {
        format!(
            "{}\u{1f}{}\u{1f}{}",
            self.cn.to_lowercase(),
            self.en.to_lowercase(),
            self.type_name.to_lowercase()
        )
    }

    pub fn base_key(&self) -> String {
        let en_lower = self.en.to_lowercase();
        if !PLACEHOLDER_HEADERS.contains(&en_lower.as_str()) {
            format!("en:{en_lower}")
        } else {
            format!("header:{}", self.identity())
        }
    }

    pub fn display_name(&self) -> String {
        match (self.cn.is_empty(), self.en.is_empty()) {
            (false, false) if self.cn != self.en => format!("{} ({})", self.en, self.cn),
            (false, false) => self.en.clone(),
            (true, false) => self.en.clone(),
            (false, true) => self.cn.clone(),
            (true, true) => "(unnamed)".to_string(),
        }
    }

    pub fn is_blank(&self) -> bool {
        self.cn.is_empty() && self.en.is_empty() && self.type_name.is_empty()
    }
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct Cell {
    pub value: String,
    pub formula: Option<String>,
}

impl Cell {
    pub fn is_empty(&self) -> bool {
        self.value.is_empty() && self.formula.is_none()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct Row {
    pub key: String,
    pub occurrence: usize,
    pub cells: Vec<Cell>,
    pub source_row: Option<usize>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Sheet {
    pub name: String,
    pub headers: Vec<Header>,
    pub rows: Vec<Row>,
}

impl Sheet {
    pub fn column_keys(&self) -> Vec<String> {
        let mut counts = HashMap::<String, usize>::new();
        self.headers
            .iter()
            .map(|header| {
                let base = header.base_key();
                let count = counts.entry(base.clone()).or_default();
                *count += 1;
                if *count == 1 {
                    base
                } else {
                    format!("{base}#{}", *count)
                }
            })
            .collect()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct Workbook {
    pub path: String,
    pub sheets: Vec<Sheet>,
    pub issues: Vec<Issue>,
}

impl Workbook {
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    pub fn formula_count(&self) -> usize {
        self.sheets
            .iter()
            .flat_map(|sheet| sheet.rows.iter())
            .flat_map(|row| row.cells.iter())
            .filter(|cell| cell.formula.is_some())
            .count()
    }

    pub fn add_validation_issues(&mut self) {
        let mut formula_groups: HashMap<(String, String, String, String), (usize, Issue)> =
            HashMap::new();
        for sheet in self.sheets.iter_mut() {
            validate_headers(sheet, &mut self.issues);
            validate_rows(sheet, &mut self.issues);
            for row in sheet.rows.iter_mut() {
                for (column_index, cell) in row.cells.iter_mut().enumerate() {
                    let header = sheet.headers.get(column_index).cloned().unwrap_or_default();
                    if let Some(formula) = cell.formula.as_deref() {
                        let findings = analyze_formula(formula, row.source_row.unwrap_or(4));
                        for finding in findings {
                            let key = (
                                sheet.name.clone(),
                                header.display_name(),
                                finding.code.clone(),
                                finding.severity.clone(),
                            );
                            if let Some((count, existing)) = formula_groups.get_mut(&key) {
                                *count += 1;
                                let _ = existing;
                            } else {
                                formula_groups.insert(
                                    key,
                                    (
                                        1,
                                        Issue::new(
                                            &finding.code,
                                            &finding.severity,
                                            finding.message,
                                            Some(&sheet.name),
                                            row.source_row,
                                            Some(&header.display_name()),
                                            finding.reference,
                                        ),
                                    ),
                                );
                            }
                        }
                    }
                    if is_excel_error(&cell.value) {
                        self.issues.push(Issue::new(
                            "excel_error_value",
                            "error",
                            format!("检测到 Excel 错误值: {}", cell.value),
                            Some(&sheet.name),
                            row.source_row,
                            Some(&header.display_name()),
                            Some(cell.value.clone()),
                        ));
                    }
                }
            }
        }
        for (_, (count, mut issue)) in formula_groups {
            if count > 1 {
                issue
                    .message
                    .push_str(&format!(" 共 {count} 个单元格触发。"));
            }
            self.issues.push(issue);
        }
    }
}

fn validate_headers(sheet: &Sheet, issues: &mut Vec<Issue>) {
    let mut cn_map = HashMap::<String, Vec<usize>>::new();
    let mut en_map = HashMap::<String, Vec<usize>>::new();
    let mut triple_map = HashMap::<String, Vec<usize>>::new();
    for (index, header) in sheet.headers.iter().enumerate() {
        if !header.cn.is_empty() {
            cn_map
                .entry(header.cn.to_lowercase())
                .or_default()
                .push(index);
        }
        if !PLACEHOLDER_HEADERS.contains(&header.en.to_lowercase().as_str()) {
            en_map
                .entry(header.en.to_lowercase())
                .or_default()
                .push(index);
        }
        if !header.is_blank() {
            triple_map.entry(header.identity()).or_default().push(index);
        }
        if header.en.is_empty()
            && sheet
                .rows
                .iter()
                .any(|row| row.cells.get(index).is_some_and(|cell| !cell.is_empty()))
        {
            issues.push(Issue::new(
                "missing_english_header",
                "warning",
                "该列有数据，但缺少英文字段名。",
                Some(&sheet.name),
                None,
                Some(&column_letter(index + 1)),
                None,
            ));
        }
    }
    for (code, severity, label, mapping) in [
        ("duplicate_english_header", "error", "英文字段名", en_map),
        ("duplicate_chinese_header", "warning", "中文字段名", cn_map),
        ("duplicate_header", "warning", "完整表头", triple_map),
    ] {
        for (name, indexes) in mapping {
            if indexes.len() > 1 {
                issues.push(Issue::new(
                    code,
                    severity,
                    format!(
                        "{label}重复；列: {}",
                        indexes
                            .iter()
                            .map(|index| column_letter(index + 1))
                            .collect::<Vec<_>>()
                            .join(", ")
                    ),
                    Some(&sheet.name),
                    None,
                    Some(&column_letter(indexes[0] + 1)),
                    Some(name),
                ));
            }
        }
    }
}

fn validate_rows(sheet: &Sheet, issues: &mut Vec<Issue>) {
    let mut keys = HashMap::<String, Vec<usize>>::new();
    for row in &sheet.rows {
        keys.entry(row.key.clone())
            .or_default()
            .push(row.source_row.unwrap_or(0));
    }
    for (key, rows) in keys {
        if rows.len() > 1 {
            issues.push(Issue::new(
                "duplicate_row_key",
                "error",
                "重复 RowName；所有重复行都保留在模型中，但三路合并不会静默覆盖。",
                Some(&sheet.name),
                rows.first().copied(),
                None,
                Some(key),
            ));
        }
    }
}

fn is_excel_error(value: &str) -> bool {
    matches!(
        value.trim(),
        "#N/A" | "#REF!" | "#VALUE!" | "#NAME?" | "#DIV/0!" | "#NUM!" | "#NULL!"
    )
}

fn value_to_text(value: &Data) -> String {
    match value {
        Data::Int(value) => value.to_string(),
        Data::Float(value) => {
            if value.is_finite() && value.fract() == 0.0 {
                format!("{}", *value as i64)
            } else {
                value.to_string()
            }
        }
        Data::String(value) => value.trim().to_string(),
        Data::Bool(value) => if *value { "TRUE" } else { "FALSE" }.to_string(),
        Data::DateTime(value) => value.to_string(),
        Data::DateTimeIso(value) | Data::DurationIso(value) => value.trim().to_string(),
        Data::Error(value) => value.to_string(),
        Data::Empty => String::new(),
    }
}

pub fn read_workbook(path: impl AsRef<Path>) -> Result<Workbook> {
    let path = path.as_ref();
    let mut workbook: Sheets<_> = open_workbook_auto(path).map_err(CoreError::from)?;
    let sheet_names = workbook.sheet_names();
    let mut result = Workbook {
        path: path.to_string_lossy().to_string(),
        sheets: Vec::new(),
        issues: Vec::new(),
    };
    for sheet_name in sheet_names.into_iter().filter(|name| name.starts_with('#')) {
        let values = workbook.worksheet_range(&sheet_name)?;
        let formulas = workbook.worksheet_formula(&sheet_name).unwrap_or_default();
        let start = values.start().unwrap_or((0, 0));
        let end = values.end().unwrap_or((0, 0));
        if values.is_empty() {
            continue;
        }
        let used_columns: Vec<u32> = (start.1..=end.1)
            .filter(|column| {
                (start.0..=end.0).any(|row| {
                    values
                        .get_value((row, *column))
                        .is_some_and(|value| !value_to_text(value).is_empty())
                        || formulas
                            .get_value((row, *column))
                            .is_some_and(|value| !value.trim().is_empty())
                })
            })
            .collect();
        let header = |row_index: u32| {
            used_columns
                .iter()
                .map(|column| {
                    values
                        .get_value((row_index, *column))
                        .map(value_to_text)
                        .unwrap_or_default()
                })
                .collect::<Vec<_>>()
        };
        let rows0 = header(0);
        let rows1 = header(1);
        let rows2 = header(2);
        let headers = (0..used_columns.len())
            .map(|index| {
                Header::new(
                    rows0.get(index).cloned().unwrap_or_default(),
                    rows1.get(index).cloned().unwrap_or_default(),
                    rows2.get(index).cloned().unwrap_or_default(),
                )
            })
            .collect::<Vec<_>>();
        let mut rows = Vec::new();
        let mut occurrences = HashMap::<String, usize>::new();
        for row_index in start.0..=end.0 {
            if row_index < (DATA_START_ROW - 1) as u32 {
                continue;
            }
            let cells = used_columns
                .iter()
                .map(|column| {
                    let value = values
                        .get_value((row_index, *column))
                        .map(value_to_text)
                        .unwrap_or_default();
                    let formula = formulas
                        .get_value((row_index, *column))
                        .map(|value| value.trim().to_string())
                        .filter(|value| !value.is_empty());
                    Cell { value, formula }
                })
                .collect::<Vec<_>>();
            if cells.iter().all(Cell::is_empty) {
                if row_index + 1 < end.0 {
                    result.issues.push(Issue::new(
                        "blank_row_removed",
                        "warning",
                        "检测到空数据行，加载时已从模型中移除。",
                        Some(&sheet_name),
                        Some(row_index as usize + 1),
                        None,
                        None,
                    ));
                }
                continue;
            }
            let mut key = cells
                .first()
                .map(|cell| cell.value.clone())
                .unwrap_or_default();
            if key.is_empty()
                && let Some(formula) = cells.first().and_then(|cell| cell.formula.clone())
            {
                key = cells
                    .get(1)
                    .filter(|cell| !cell.value.is_empty())
                    .map(|cell| cell.value.clone())
                    .unwrap_or(formula);
            }
            if key.trim().is_empty() {
                key = format!("__row_{}", row_index + 1);
                result.issues.push(Issue::new(
                    "missing_row_key",
                    "warning",
                    "数据行没有 RowName；已使用临时行键，保存/合并前应补全。",
                    Some(&sheet_name),
                    Some(row_index as usize + 1),
                    None,
                    None,
                ));
            }
            let key = key.trim().to_string();
            let occurrence = occurrences.entry(key.clone()).or_default();
            *occurrence += 1;
            rows.push(Row {
                key,
                occurrence: *occurrence,
                cells,
                source_row: Some(row_index as usize + 1),
            });
        }
        result.sheets.push(Sheet {
            name: sheet_name,
            headers,
            rows,
        });
    }
    Ok(result)
}

pub fn column_letter(mut column: usize) -> String {
    let mut result = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        result.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    result
}
