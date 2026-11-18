use std::collections::{HashMap, HashSet};
use std::path::Path;

use serde::Serialize;

use crate::error::Result;
use crate::model::{Header, Issue, Workbook, read_workbook};
use crate::snapshot::{TextSheet, workbook_to_text};

#[derive(Debug, Clone, Serialize)]
pub struct CellChange {
    pub column: String,
    pub old: String,
    pub new: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct RowChange {
    pub sheet: String,
    pub row: String,
    pub kind: String,
    pub cells: Vec<CellChange>,
    pub old_values: Vec<String>,
    pub new_values: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ColumnChange {
    pub sheet: String,
    pub column: String,
    pub kind: String,
    pub old: Option<Header>,
    pub new: Option<Header>,
    pub old_index: Option<usize>,
    pub new_index: Option<usize>,
}

#[derive(Debug, Clone, Serialize)]
pub struct ColumnAlignment {
    pub kind: String,
    pub old_index: Option<usize>,
    pub new_index: Option<usize>,
    pub old_header: Option<Header>,
    pub new_header: Option<Header>,
}

#[derive(Debug, Clone, Serialize)]
pub struct SheetDiff {
    pub sheet: String,
    pub status: String,
    pub columns: Vec<Header>,
    pub column_alignments: Vec<ColumnAlignment>,
    pub column_changes: Vec<ColumnChange>,
    pub row_changes: Vec<RowChange>,
}

impl SheetDiff {
    pub fn has_changes(&self) -> bool {
        !self.column_changes.is_empty() || !self.row_changes.is_empty()
    }
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct DiffSummary {
    pub rows_added: usize,
    pub rows_deleted: usize,
    pub rows_modified: usize,
    pub columns_added: usize,
    pub columns_deleted: usize,
    pub columns_modified: usize,
    pub columns_moved: usize,
    pub cells_modified: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct DiffResult {
    pub old: String,
    pub new: String,
    pub comparison: String,
    pub has_changes: bool,
    pub summary: DiffSummary,
    pub sheets: Vec<SheetDiff>,
    pub issues: Vec<Issue>,
}

#[derive(Debug, Clone)]
struct AlignedColumn {
    key: String,
    kind: String,
    old_index: Option<usize>,
    new_index: Option<usize>,
    old_header: Option<Header>,
    new_header: Option<Header>,
}

pub fn diff_files(old_path: impl AsRef<Path>, new_path: impl AsRef<Path>) -> Result<DiffResult> {
    let mut old = read_workbook(old_path)?;
    let mut new = read_workbook(new_path)?;
    old.add_validation_issues();
    new.add_validation_issues();
    Ok(diff_workbooks(&old, &new))
}

pub fn diff_workbooks(old: &Workbook, new: &Workbook) -> DiffResult {
    let old_snapshot = workbook_to_text(old);
    let new_snapshot = workbook_to_text(new);
    let mut sheet_names: Vec<String> = Vec::new();
    for sheet in old_snapshot.sheets.iter().chain(new_snapshot.sheets.iter()) {
        if !sheet_names.contains(&sheet.name) {
            sheet_names.push(sheet.name.clone());
        }
    }
    let mut sheets = Vec::new();
    for sheet_name in sheet_names {
        let old_sheet = old_snapshot
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name);
        let new_sheet = new_snapshot
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name);
        match (old_sheet, new_sheet) {
            (None, Some(new_sheet)) => sheets.push(SheetDiff {
                sheet: sheet_name,
                status: "added".to_string(),
                columns: new_sheet.headers(),
                column_alignments: new_sheet
                    .columns
                    .iter()
                    .enumerate()
                    .map(|(index, column)| ColumnAlignment {
                        kind: "added".to_string(),
                        old_index: None,
                        new_index: Some(index),
                        old_header: None,
                        new_header: Some(column.header.clone()),
                    })
                    .collect(),
                column_changes: vec![],
                row_changes: new_sheet
                    .rows
                    .iter()
                    .map(|row| RowChange {
                        sheet: new_sheet.name.clone(),
                        row: row.key.clone(),
                        kind: "added".to_string(),
                        cells: vec![],
                        old_values: vec![],
                        new_values: row.cells.clone(),
                    })
                    .collect(),
            }),
            (Some(old_sheet), None) => sheets.push(SheetDiff {
                sheet: sheet_name,
                status: "deleted".to_string(),
                columns: old_sheet.headers(),
                column_alignments: old_sheet
                    .columns
                    .iter()
                    .enumerate()
                    .map(|(index, column)| ColumnAlignment {
                        kind: "deleted".to_string(),
                        old_index: Some(index),
                        new_index: None,
                        old_header: Some(column.header.clone()),
                        new_header: None,
                    })
                    .collect(),
                column_changes: vec![],
                row_changes: old_sheet
                    .rows
                    .iter()
                    .map(|row| RowChange {
                        sheet: old_sheet.name.clone(),
                        row: row.key.clone(),
                        kind: "deleted".to_string(),
                        cells: vec![],
                        old_values: row.cells.clone(),
                        new_values: vec![],
                    })
                    .collect(),
            }),
            (Some(old_sheet), Some(new_sheet)) => {
                sheets.push(diff_sheet(old_sheet, new_sheet));
            }
            (None, None) => {}
        }
    }

    let mut issues = old.issues.clone();
    issues.extend(new.issues.clone());
    issues.extend(old_snapshot.issues.clone());
    issues.extend(new_snapshot.issues.clone());
    let summary = summarize(&sheets);
    DiffResult {
        old: old.path.clone(),
        new: new.path.clone(),
        comparison: "text_snapshot_with_cached_formula_values".to_string(),
        has_changes: sheets.iter().any(SheetDiff::has_changes),
        summary,
        sheets,
        issues,
    }
}

fn summarize(sheets: &[SheetDiff]) -> DiffSummary {
    let mut summary = DiffSummary::default();
    for sheet in sheets {
        for change in &sheet.column_changes {
            match change.kind.as_str() {
                "added" => summary.columns_added += 1,
                "deleted" => summary.columns_deleted += 1,
                "modified" => summary.columns_modified += 1,
                "moved" => summary.columns_moved += 1,
                _ => {}
            }
        }
        for change in &sheet.row_changes {
            match change.kind.as_str() {
                "added" => summary.rows_added += 1,
                "deleted" => summary.rows_deleted += 1,
                "modified" => summary.rows_modified += 1,
                _ => {}
            }
            summary.cells_modified += change.cells.len();
        }
    }
    summary
}

fn diff_sheet(old: &TextSheet, new: &TextSheet) -> SheetDiff {
    let (aligned, column_changes) = align_columns(old, new);
    let column_alignments = aligned
        .iter()
        .map(|column| ColumnAlignment {
            kind: column.kind.clone(),
            old_index: column.old_index,
            new_index: column.new_index,
            old_header: column.old_header.clone(),
            new_header: column.new_header.clone(),
        })
        .collect();
    let row_changes = diff_rows(old, new, &aligned);
    SheetDiff {
        sheet: new.name.clone(),
        status: if column_changes.is_empty() && row_changes.is_empty() {
            "same".to_string()
        } else {
            "modified".to_string()
        },
        columns: new.headers(),
        column_alignments,
        column_changes,
        row_changes,
    }
}

fn align_columns(old: &TextSheet, new: &TextSheet) -> (Vec<AlignedColumn>, Vec<ColumnChange>) {
    let old_keys = old.column_keys();
    let new_keys = new.column_keys();
    let old_pos: HashMap<&str, usize> = old_keys
        .iter()
        .enumerate()
        .map(|(index, key)| (key.as_str(), index))
        .collect();
    let new_pos: HashMap<&str, usize> = new_keys
        .iter()
        .enumerate()
        .map(|(index, key)| (key.as_str(), index))
        .collect();
    let common: HashSet<&str> = old_keys
        .iter()
        .filter(|key| new_pos.contains_key(key.as_str()))
        .map(String::as_str)
        .collect();
    let mut used_old = HashSet::new();
    let mut used_new = HashSet::new();
    let mut aligned = Vec::new();
    for key in &old_keys {
        if let Some(new_index) = new_pos.get(key.as_str()).copied() {
            let old_index = old_pos[key.as_str()];
            aligned.push(AlignedColumn {
                key: key.clone(),
                kind: "same".to_string(),
                old_index: Some(old_index),
                new_index: Some(new_index),
                old_header: Some(old.columns[old_index].header.clone()),
                new_header: Some(new.columns[new_index].header.clone()),
            });
            used_old.insert(old_index);
            used_new.insert(new_index);
        }
    }

    let unmatched_old: Vec<usize> = (0..old.columns.len())
        .filter(|index| !used_old.contains(index))
        .collect();
    let unmatched_new: Vec<usize> = (0..new.columns.len())
        .filter(|index| !used_new.contains(index))
        .collect();
    let mut changes = Vec::new();
    let pairs = unmatched_old.len().min(unmatched_new.len());
    for pair in 0..pairs {
        let old_index = unmatched_old[pair];
        let new_index = unmatched_new[pair];
        let old_header = old.columns[old_index].header.clone();
        let new_header = new.columns[new_index].header.clone();
        aligned.push(AlignedColumn {
            key: old_keys[old_index].clone(),
            kind: "modified".to_string(),
            old_index: Some(old_index),
            new_index: Some(new_index),
            old_header: Some(old_header.clone()),
            new_header: Some(new_header.clone()),
        });
        changes.push(ColumnChange {
            sheet: new.name.clone(),
            column: new_header.display_name(),
            kind: "modified".to_string(),
            old: Some(old_header),
            new: Some(new_header),
            old_index: Some(old_index),
            new_index: Some(new_index),
        });
    }
    for old_index in &unmatched_old[pairs..] {
        let old_header = old.columns[*old_index].header.clone();
        changes.push(ColumnChange {
            sheet: new.name.clone(),
            column: old_header.display_name(),
            kind: "deleted".to_string(),
            old: Some(old_header),
            new: None,
            old_index: Some(*old_index),
            new_index: None,
        });
    }
    for new_index in &unmatched_new[pairs..] {
        let new_header = new.columns[*new_index].header.clone();
        changes.push(ColumnChange {
            sheet: new.name.clone(),
            column: new_header.display_name(),
            kind: "added".to_string(),
            old: None,
            new: Some(new_header),
            old_index: None,
            new_index: Some(*new_index),
        });
    }

    let common_old: Vec<String> = old_keys
        .iter()
        .filter(|key| common.contains(key.as_str()))
        .cloned()
        .collect();
    let common_new: Vec<String> = new_keys
        .iter()
        .filter(|key| common.contains(key.as_str()))
        .cloned()
        .collect();
    let keep = lcs_keep(&common_old, &common_new);
    let moved: HashSet<&str> = common
        .iter()
        .enumerate()
        .filter(|(index, _)| !keep.0.contains(index))
        .map(|(_, key)| *key)
        .collect();
    for item in &aligned {
        if item.kind == "same" && moved.contains(item.key.as_str()) {
            changes.push(ColumnChange {
                sheet: new.name.clone(),
                column: item
                    .new_header
                    .as_ref()
                    .map(Header::display_name)
                    .unwrap_or_else(|| item.key.clone()),
                kind: "moved".to_string(),
                old: item.old_header.clone(),
                new: item.new_header.clone(),
                old_index: item.old_index,
                new_index: item.new_index,
            });
        }
    }
    aligned.sort_by_key(|item| {
        (
            item.new_index.is_none(),
            item.new_index.or(item.old_index).unwrap_or(usize::MAX),
        )
    });
    changes.sort_by_key(|change| {
        (
            change
                .new_index
                .unwrap_or(10_000 + change.old_index.unwrap_or(0)),
            change.kind.clone(),
        )
    });
    (aligned, changes)
}

fn lcs_keep(old: &[String], new: &[String]) -> (HashSet<usize>, HashSet<usize>) {
    let mut lengths = vec![vec![0usize; new.len() + 1]; old.len() + 1];
    for i in (0..old.len()).rev() {
        for j in (0..new.len()).rev() {
            lengths[i][j] = if old[i] == new[j] {
                lengths[i + 1][j + 1] + 1
            } else {
                lengths[i + 1][j].max(lengths[i][j + 1])
            };
        }
    }
    let mut old_keep = HashSet::new();
    let mut new_keep = HashSet::new();
    let (mut i, mut j) = (0, 0);
    while i < old.len() && j < new.len() {
        if old[i] == new[j] {
            old_keep.insert(i);
            new_keep.insert(j);
            i += 1;
            j += 1;
        } else if lengths[i + 1][j] >= lengths[i][j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    (old_keep, new_keep)
}

fn row_map(sheet: &TextSheet) -> HashMap<(String, usize), usize> {
    sheet
        .rows
        .iter()
        .enumerate()
        .map(|(index, row)| ((row.key.clone(), row.occurrence), index))
        .collect()
}

fn diff_rows(old: &TextSheet, new: &TextSheet, aligned: &[AlignedColumn]) -> Vec<RowChange> {
    let old_map = row_map(old);
    let new_map = row_map(new);
    let mut identities: Vec<(String, usize)> = Vec::new();
    for row in &old.rows {
        identities.push((row.key.clone(), row.occurrence));
    }
    for row in &new.rows {
        let identity = (row.key.clone(), row.occurrence);
        if !identities.contains(&identity) {
            identities.push(identity);
        }
    }
    let modified_columns: HashSet<&str> = aligned
        .iter()
        .filter(|item| item.kind == "modified")
        .map(|item| item.key.as_str())
        .collect();
    let compared: Vec<&AlignedColumn> = aligned
        .iter()
        .filter(|item| {
            item.old_index.is_some()
                && item.new_index.is_some()
                && !modified_columns.contains(item.key.as_str())
        })
        .collect();
    let mut changes = Vec::new();
    for identity in identities {
        let old_index = old_map.get(&identity).copied();
        let new_index = new_map.get(&identity).copied();
        match (old_index, new_index) {
            (None, Some(new_index)) => changes.push(RowChange {
                sheet: new.name.clone(),
                row: identity.0,
                kind: "added".to_string(),
                cells: vec![],
                old_values: vec![],
                new_values: new.rows[new_index].cells.clone(),
            }),
            (Some(old_index), None) => changes.push(RowChange {
                sheet: old.name.clone(),
                row: identity.0,
                kind: "deleted".to_string(),
                cells: vec![],
                old_values: old.rows[old_index].cells.clone(),
                new_values: vec![],
            }),
            (Some(old_index), Some(new_index)) => {
                let old_row = &old.rows[old_index];
                let new_row = &new.rows[new_index];
                let cells = compared
                    .iter()
                    .filter_map(|column| {
                        let old_cell = old_row.cells.get(column.old_index?)?;
                        let new_cell = new_row.cells.get(column.new_index?)?;
                        if old_cell == new_cell {
                            None
                        } else {
                            Some(CellChange {
                                column: column
                                    .new_header
                                    .as_ref()
                                    .or(column.old_header.as_ref())
                                    .map(Header::display_name)
                                    .unwrap_or_default(),
                                old: old_cell.clone(),
                                new: new_cell.clone(),
                            })
                        }
                    })
                    .collect::<Vec<_>>();
                if !cells.is_empty() {
                    changes.push(RowChange {
                        sheet: new.name.clone(),
                        row: identity.0,
                        kind: "modified".to_string(),
                        cells,
                        old_values: old_row.cells.clone(),
                        new_values: new_row.cells.clone(),
                    });
                }
            }
            (None, None) => {}
        }
    }
    changes
}

pub fn format_text_diff(result: &DiffResult) -> String {
    let old_name = Path::new(&result.old)
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("old");
    let new_name = Path::new(&result.new)
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("new");
    let mut lines = Vec::new();
    for sheet in result.sheets.iter().filter(|sheet| sheet.has_changes()) {
        lines.push(format!("diff --table {}", sheet.sheet));
        lines.push(format!("--- {old_name}"));
        lines.push(format!("+++ {new_name}"));
        if !sheet.column_changes.is_empty() {
            lines.push("@@ columns @@".to_string());
            for change in &sheet.column_changes {
                match change.kind.as_str() {
                    "added" => lines.push(format!("+ column {}", change.column)),
                    "deleted" => lines.push(format!("- column {}", change.column)),
                    "modified" => lines.push(format!("~ column {}", change.column)),
                    "moved" => lines.push(format!("~ column moved {}", change.column)),
                    _ => {}
                }
            }
        }
        for change in &sheet.row_changes {
            lines.push(format!("@@ row {} @@", change.row));
            match change.kind.as_str() {
                "added" => {
                    lines.push(format!("+ row {}", change.row));
                    for (index, value) in change.new_values.iter().enumerate() {
                        lines.push(format!("+   col{}: {value}", index + 1));
                    }
                }
                "deleted" => {
                    lines.push(format!("- row {}", change.row));
                    for (index, value) in change.old_values.iter().enumerate() {
                        lines.push(format!("-   col{}: {value}", index + 1));
                    }
                }
                "modified" => {
                    for cell in &change.cells {
                        lines.push(format!("- {}: {}", cell.column, cell.old));
                        lines.push(format!("+ {}: {}", cell.column, cell.new));
                    }
                }
                _ => {}
            }
        }
        lines.push(String::new());
    }
    if lines.is_empty() {
        lines.push("No differences.".to_string());
    }
    if !result.issues.is_empty() {
        lines.push("Warnings:".to_string());
        for issue in &result.issues {
            lines.push(format!(
                "- [{}] {}: {}",
                issue.severity, issue.code, issue.message
            ));
        }
    }
    lines.join("\n")
}
