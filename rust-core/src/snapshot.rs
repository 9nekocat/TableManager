use serde::Serialize;
use std::collections::HashMap;

use crate::model::{Header, Issue, Workbook};

#[derive(Debug, Clone, Serialize)]
pub struct TextColumn {
    pub key: String,
    pub header: Header,
}

#[derive(Debug, Clone, Serialize)]
pub struct TextRow {
    pub key: String,
    pub occurrence: usize,
    pub cells: Vec<String>,
    pub source_row: Option<usize>,
}

#[derive(Debug, Clone, Serialize)]
pub struct TextSheet {
    pub name: String,
    pub columns: Vec<TextColumn>,
    pub rows: Vec<TextRow>,
}

impl TextSheet {
    pub fn headers(&self) -> Vec<Header> {
        self.columns
            .iter()
            .map(|column| column.header.clone())
            .collect()
    }

    pub fn column_keys(&self) -> Vec<String> {
        self.columns
            .iter()
            .map(|column| column.key.clone())
            .collect()
    }
}

#[derive(Debug, Clone, Serialize)]
pub struct TextWorkbook {
    pub path: String,
    pub sheets: Vec<TextSheet>,
    pub issues: Vec<Issue>,
}

pub fn workbook_to_text(workbook: &Workbook) -> TextWorkbook {
    let mut issues = Vec::new();
    let mut sheets = Vec::new();
    for sheet in &workbook.sheets {
        let keys = sheet.column_keys();
        let columns = sheet
            .headers
            .iter()
            .cloned()
            .zip(keys)
            .map(|(header, key)| TextColumn { key, header })
            .collect::<Vec<_>>();
        let mut missing_cache: HashMap<String, (usize, String, Option<usize>)> = HashMap::new();
        let mut rows = Vec::new();
        for row in &sheet.rows {
            let mut cells = Vec::with_capacity(row.cells.len());
            for (index, cell) in row.cells.iter().enumerate() {
                if cell.formula.is_some() && cell.value.is_empty() {
                    let column = sheet
                        .headers
                        .get(index)
                        .map(Header::display_name)
                        .unwrap_or_default();
                    let entry = missing_cache.entry(column).or_insert((
                        0,
                        cell.formula.clone().unwrap_or_default(),
                        row.source_row,
                    ));
                    entry.0 += 1;
                }
                cells.push(cell.value.clone());
            }
            rows.push(TextRow {
                key: row.key.clone(),
                occurrence: row.occurrence,
                cells,
                source_row: row.source_row,
            });
        }
        for (column, (count, formula, row)) in missing_cache {
            issues.push(Issue::new(
                "formula_value_unavailable",
                "warning",
                if count > 1 {
                    format!("公式没有可用的 Excel 缓存值，差异中按空值处理。 共 {count} 个单元格。")
                } else {
                    "公式没有可用的 Excel 缓存值，差异中按空值处理。".to_string()
                },
                Some(&sheet.name),
                row,
                Some(&column),
                Some(formula),
            ));
        }
        sheets.push(TextSheet {
            name: sheet.name.clone(),
            columns,
            rows,
        });
    }
    TextWorkbook {
        path: workbook.path.clone(),
        sheets,
        issues,
    }
}

pub fn render_text_workbook(snapshot: &TextWorkbook) -> String {
    let mut lines = Vec::new();
    for sheet in &snapshot.sheets {
        lines.push(format!("## {}", sheet.name));
        lines.push(format!(
            "columns\t{}",
            sheet
                .columns
                .iter()
                .map(|column| format!(
                    "{}|{}|{}",
                    column.header.cn, column.header.en, column.header.type_name
                ))
                .collect::<Vec<_>>()
                .join("\t")
        ));
        for row in &sheet.rows {
            lines.push(format!(
                "row\t{}#{}\t{}",
                row.key,
                row.occurrence,
                row.cells
                    .iter()
                    .map(|cell| cell.replace('\t', "\\t").replace('\n', "\\n"))
                    .collect::<Vec<_>>()
                    .join("\t")
            ));
        }
    }
    lines.join("\n")
}
