use std::collections::HashMap;
use std::path::Path;

use calamine::{Data, Range, Reader, Sheets, open_workbook_auto};

use crate::error::{CoreError, Result};
use crate::merge::{
    ConflictDefault, MergePreviewRow, MergeStrategy, merge_files_with_conflict_default,
};
use crate::model::{Cell, DATA_START_ROW, Row};
use crate::output::write_merged_workbook;

const META_SHEET: &str = "_MergeMeta";

#[derive(Debug)]
struct ReviewMeta {
    kind: String,
    preview_sheet: String,
    source_sheet: String,
    row: String,
    occurrence: usize,
    preview_row: u32,
    original_final: Vec<Cell>,
}

struct PreviewSheet {
    values: Range<Data>,
    formulas: Range<String>,
}

pub fn apply_merge_review(
    preview_path: impl AsRef<Path>,
    base_path: impl AsRef<Path>,
    ours_path: impl AsRef<Path>,
    theirs_path: impl AsRef<Path>,
    strategy: MergeStrategy,
    conflict_default: ConflictDefault,
    output_path: impl AsRef<Path>,
) -> Result<usize> {
    let metas = read_metadata(preview_path.as_ref())?.ok_or_else(|| {
        CoreError::Message("合并预览缺少 _MergeMeta，无法确认最终值。".to_string())
    })?;
    let previews = read_preview_sheets(preview_path.as_ref(), &metas)?;
    let mut result = merge_files_with_conflict_default(
        base_path,
        ours_path,
        theirs_path,
        strategy,
        conflict_default,
    )?;

    for meta in &metas {
        let preview = previews.get(&meta.preview_sheet).ok_or_else(|| {
            CoreError::Message(format!("合并预览缺少工作表 {}。", meta.preview_sheet))
        })?;
        let preview_row = result
            .preview_rows
            .iter()
            .find(|row| {
                row.sheet == meta.source_sheet
                    && row.row == meta.row
                    && row.occurrence == meta.occurrence
            })
            .ok_or_else(|| {
                CoreError::Message(format!(
                    "重新计算三路合并时找不到预览行 {}!{}#{}。",
                    meta.source_sheet, meta.row, meta.occurrence
                ))
            })?;
        let current_cells =
            read_preview_row(preview, meta.preview_row, preview_row.final_cells.len());
        let decision = read_preview_text(preview, meta.preview_row, 1);
        apply_review_row(
            &mut result.workbook,
            meta,
            preview_row,
            &current_cells,
            &decision,
        )?;
    }

    write_merged_workbook(&result, output_path)?;
    Ok(metas.len())
}

fn read_metadata(path: &Path) -> Result<Option<Vec<ReviewMeta>>> {
    let mut workbook: Sheets<_> = open_workbook_auto(path)?;
    if !workbook.sheet_names().iter().any(|name| name == META_SHEET) {
        return Ok(None);
    }
    let range = workbook.worksheet_range(META_SHEET)?;
    let end = range.end().unwrap_or((0, 0));
    let mut rows = Vec::new();
    for row in 1..=end.0 {
        let kind = range_text(&range, row, 0);
        if kind.is_empty() {
            continue;
        }
        let original_json = range_text(&range, row, 7);
        let original_final =
            serde_json::from_str::<Vec<Cell>>(&original_json).map_err(|error| {
                CoreError::Message(format!(
                    "无法读取合并预览第 {} 行的原始最终值: {error}",
                    row + 1
                ))
            })?;
        rows.push(ReviewMeta {
            kind,
            preview_sheet: range_text(&range, row, 1),
            source_sheet: range_text(&range, row, 2),
            row: range_text(&range, row, 3),
            occurrence: range_text(&range, row, 4).parse().map_err(|error| {
                CoreError::Message(format!(
                    "无法读取合并预览第 {} 行的行序号: {error}",
                    row + 1
                ))
            })?,
            preview_row: range_text(&range, row, 5).parse().map_err(|error| {
                CoreError::Message(format!(
                    "无法读取合并预览第 {} 行的 Excel 行号: {error}",
                    row + 1
                ))
            })?,
            original_final,
        });
    }
    Ok(Some(rows))
}

fn read_preview_sheets(path: &Path, metas: &[ReviewMeta]) -> Result<HashMap<String, PreviewSheet>> {
    let mut workbook: Sheets<_> = open_workbook_auto(path)?;
    let available = workbook.sheet_names();
    let mut result = HashMap::new();
    for meta in metas {
        if result.contains_key(&meta.preview_sheet) {
            continue;
        }
        if !available.iter().any(|name| name == &meta.preview_sheet) {
            return Err(CoreError::Message(format!(
                "合并预览缺少工作表 {}。",
                meta.preview_sheet
            )));
        }
        let values = workbook.worksheet_range(&meta.preview_sheet)?;
        let formulas = workbook
            .worksheet_formula(&meta.preview_sheet)
            .unwrap_or_default();
        result.insert(
            meta.preview_sheet.clone(),
            PreviewSheet { values, formulas },
        );
    }
    Ok(result)
}

fn apply_review_row(
    workbook: &mut crate::model::Workbook,
    meta: &ReviewMeta,
    preview_row: &MergePreviewRow,
    current_cells: &[Cell],
    decision: &str,
) -> Result<()> {
    let sheet = workbook
        .sheets
        .iter_mut()
        .find(|sheet| sheet.name == meta.source_sheet)
        .ok_or_else(|| {
            CoreError::Message(format!("最终合并结果缺少工作表 {}。", meta.source_sheet))
        })?;
    let row_index = sheet
        .rows
        .iter()
        .position(|row| row.key == meta.row && row.occurrence == meta.occurrence);
    let normalized = normalize_decision(decision);

    if meta.kind == "row_delete_conflict" {
        let delete = match normalized.as_str() {
            "删除" | "delete" => true,
            "修改" | "keep" | "保留" => false,
            "" => !preview_row.default_keep,
            _ => {
                return Err(CoreError::Message(format!(
                    "{}!{} 的行保留选项无效: {decision}",
                    meta.source_sheet, meta.row
                )));
            }
        };
        if delete {
            if let Some(index) = row_index {
                sheet.rows.remove(index);
            }
            return Ok(());
        }
        let mut cells = preview_row.final_cells.clone();
        merge_edited_cells(&mut cells, current_cells, &meta.original_final);
        if let Some(index) = row_index {
            sheet.rows[index].cells = cells;
        } else {
            let index = preview_row.order_index.min(sheet.rows.len());
            sheet.rows.insert(
                index,
                Row {
                    key: meta.row.clone(),
                    occurrence: meta.occurrence,
                    cells,
                    source_row: Some(DATA_START_ROW + index),
                },
            );
            for (offset, row) in sheet.rows.iter_mut().enumerate().skip(index) {
                row.source_row = Some(DATA_START_ROW + offset);
            }
        }
        return Ok(());
    }

    let index = row_index.ok_or_else(|| {
        CoreError::Message(format!(
            "最终合并结果缺少冲突行 {}!{}。",
            meta.source_sheet, meta.row
        ))
    })?;
    let mut cells = sheet.rows[index].cells.clone();
    let edited = edited_columns(current_cells, &meta.original_final);
    for (column, is_edited) in edited.iter().enumerate() {
        if *is_edited && let Some(cell) = current_cells.get(column) {
            cells[column] = cell.clone();
        }
    }
    match normalized.as_str() {
        "我的" | "mine" | "ours" => {
            copy_unedited_conflict_cells(
                &mut cells,
                &preview_row.ours,
                &preview_row.conflict_columns,
                &edited,
            );
        }
        "仓库" | "theirs" => {
            copy_unedited_conflict_cells(
                &mut cells,
                &preview_row.theirs,
                &preview_row.conflict_columns,
                &edited,
            );
        }
        "空" | "empty" | "空白" => {
            for column in &preview_row.conflict_columns {
                if !edited.get(*column).copied().unwrap_or(false)
                    && let Some(cell) = cells.get_mut(*column)
                {
                    *cell = Cell::default();
                }
            }
        }
        "自定义" | "custom" | "" => {}
        _ => {
            return Err(CoreError::Message(format!(
                "{}!{} 的冲突处理方式无效: {decision}",
                meta.source_sheet, meta.row
            )));
        }
    }
    sheet.rows[index].cells = cells;
    Ok(())
}

fn copy_unedited_conflict_cells(
    target: &mut [Cell],
    source: &[Cell],
    columns: &[usize],
    edited: &[bool],
) {
    for column in columns {
        if edited.get(*column).copied().unwrap_or(false) {
            continue;
        }
        if let (Some(target), Some(source)) = (target.get_mut(*column), source.get(*column)) {
            *target = source.clone();
        }
    }
}

fn edited_columns(current: &[Cell], original: &[Cell]) -> Vec<bool> {
    (0..current.len())
        .map(|index| {
            original.get(index).is_none_or(|original| {
                preview_display(original) != preview_display(&current[index])
            })
        })
        .collect()
}

fn merge_edited_cells(target: &mut [Cell], current: &[Cell], original: &[Cell]) {
    let edited = edited_columns(current, original);
    for (index, cell) in current.iter().enumerate() {
        if edited[index]
            && let Some(target) = target.get_mut(index)
        {
            *target = cell.clone();
        }
    }
}

fn normalize_decision(value: &str) -> String {
    value
        .trim()
        .trim_start_matches('[')
        .trim_end_matches(']')
        .to_lowercase()
}

fn read_preview_row(sheet: &PreviewSheet, row: u32, width: usize) -> Vec<Cell> {
    (0..width)
        .map(|column| read_preview_cell(sheet, row, column as u16 + 2))
        .collect()
}

fn read_preview_cell(sheet: &PreviewSheet, row: u32, column: u16) -> Cell {
    let value = sheet
        .values
        .get_value((row, column as u32))
        .map(data_text)
        .unwrap_or_default();
    let formula = sheet
        .formulas
        .get_value((row, column as u32))
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty());
    Cell { value, formula }
}

fn read_preview_text(sheet: &PreviewSheet, row: u32, column: u16) -> String {
    sheet
        .values
        .get_value((row, column as u32))
        .map(data_text)
        .unwrap_or_default()
}

fn range_text(range: &Range<Data>, row: u32, column: u32) -> String {
    range
        .get_value((row, column))
        .map(data_text)
        .unwrap_or_default()
}

fn data_text(value: &Data) -> String {
    match value {
        Data::Empty => String::new(),
        Data::Int(value) => value.to_string(),
        Data::Float(value) => {
            if value.is_finite() && value.fract() == 0.0 {
                format!("{value:.0}")
            } else {
                value.to_string()
            }
        }
        Data::String(value) => value.trim().to_string(),
        Data::Bool(value) => if *value { "TRUE" } else { "FALSE" }.to_string(),
        Data::DateTime(value) => value.to_string(),
        Data::DateTimeIso(value) | Data::DurationIso(value) => value.trim().to_string(),
        Data::Error(value) => value.to_string(),
    }
}

fn preview_display(cell: &Cell) -> String {
    match (cell.formula.as_deref(), cell.value.is_empty()) {
        (Some(formula), false) => format!("={formula} -> {}", cell.value),
        (Some(formula), true) => format!("={formula}"),
        (None, _) => cell.value.clone(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Header, Sheet, Workbook};

    fn row(key: &str, source_row: usize) -> Row {
        Row {
            key: key.to_string(),
            occurrence: 1,
            cells: vec![
                Cell {
                    value: key.to_string(),
                    formula: None,
                },
                Cell {
                    value: format!("{key} value"),
                    formula: None,
                },
            ],
            source_row: Some(source_row),
        }
    }

    #[test]
    fn reinserts_row_at_original_position() {
        let mut workbook = Workbook {
            path: "output".to_string(),
            sheets: vec![Sheet {
                name: "#Data".to_string(),
                headers: vec![
                    Header::new("行名", "Row", "FString"),
                    Header::new("名称", "Name", "FString"),
                ],
                rows: vec![row("A", 4), row("C", 5)],
            }],
            issues: Vec::new(),
        };
        let meta = ReviewMeta {
            kind: "row_delete_conflict".to_string(),
            preview_sheet: "#Data".to_string(),
            source_sheet: "#Data".to_string(),
            row: "B".to_string(),
            occurrence: 1,
            preview_row: 7,
            original_final: row("B", 5).cells,
        };
        let preview = MergePreviewRow {
            sheet: "#Data".to_string(),
            row: "B".to_string(),
            occurrence: 1,
            kind: "row_delete_conflict".to_string(),
            base: row("B", 5).cells.clone(),
            ours: row("B", 5).cells.clone(),
            theirs: vec![Cell::default(); 2],
            final_cells: row("B", 5).cells,
            conflict_columns: Vec::new(),
            default_decision: ConflictDefault::Theirs,
            default_keep: false,
            order_index: 1,
        };

        apply_review_row(
            &mut workbook,
            &meta,
            &preview,
            &preview.final_cells,
            "[修改]",
        )
        .unwrap();

        assert_eq!(
            workbook.sheets[0]
                .rows
                .iter()
                .map(|row| row.key.as_str())
                .collect::<Vec<_>>(),
            vec!["A", "B", "C"]
        );
    }
}
