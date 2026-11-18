use std::path::Path;

use rust_xlsxwriter::{Color, Format, Workbook as XlsxWorkbook};

use crate::diff::DiffResult;
use crate::error::Result;
use crate::merge::MergeResult;
use crate::model::DATA_START_ROW;

pub fn write_diff_excel(result: &DiffResult, output_path: impl AsRef<Path>) -> Result<()> {
    let mut workbook = XlsxWorkbook::new();
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#4472C4");
    let added_format = Format::new().set_background_color("#E2F0D9");
    let deleted_format = Format::new().set_background_color("#FCE4D6");
    let modified_format = Format::new().set_background_color("#FFF2CC");
    let summary = workbook.add_worksheet();
    summary.set_name("Summary")?;
    for (row, (label, value)) in [
        ("Old", result.old.clone()),
        ("New", result.new.clone()),
        ("Comparison", result.comparison.clone()),
        ("Has changes", result.has_changes.to_string()),
        ("Rows added", result.summary.rows_added.to_string()),
        ("Rows deleted", result.summary.rows_deleted.to_string()),
        ("Rows modified", result.summary.rows_modified.to_string()),
        ("Columns added", result.summary.columns_added.to_string()),
        (
            "Columns deleted",
            result.summary.columns_deleted.to_string(),
        ),
        (
            "Columns modified",
            result.summary.columns_modified.to_string(),
        ),
        ("Columns moved", result.summary.columns_moved.to_string()),
        ("Cells modified", result.summary.cells_modified.to_string()),
    ]
    .into_iter()
    .enumerate()
    {
        summary.write_string_with_format(row as u32, 0, label, &header_format)?;
        summary.write_string(row as u32, 1, value)?;
    }
    summary.set_column_width(0, 24)?;
    summary.set_column_width(1, 72)?;

    let changes = workbook.add_worksheet();
    changes.set_name("Changes")?;
    for (column, name) in ["Action", "Sheet", "Row", "Column", "Old", "New", "Note"]
        .into_iter()
        .enumerate()
    {
        changes.write_string_with_format(0, column as u16, name, &header_format)?;
    }
    let mut output_row = 1u32;
    for sheet in &result.sheets {
        for change in &sheet.column_changes {
            let action = match change.kind.as_str() {
                "added" => "COLUMN_ADDED",
                "deleted" => "COLUMN_DELETED",
                "modified" => "COLUMN_MODIFIED",
                "moved" => "COLUMN_MOVED",
                _ => "COLUMN_CHANGE",
            };
            let old = change
                .old
                .as_ref()
                .map(|header| header.display_name())
                .unwrap_or_default();
            let new = change
                .new
                .as_ref()
                .map(|header| header.display_name())
                .unwrap_or_default();
            write_row(
                changes,
                output_row,
                &[action, &sheet.sheet, "", &change.column, &old, &new, ""],
                format_for_kind(
                    &change.kind,
                    &added_format,
                    &deleted_format,
                    &modified_format,
                ),
            )?;
            output_row += 1;
        }
        for change in &sheet.row_changes {
            match change.kind.as_str() {
                "added" => {
                    let values = change.new_values.join(" | ");
                    write_row(
                        changes,
                        output_row,
                        &["ROW_ADDED", &sheet.sheet, &change.row, "", "", &values, ""],
                        &added_format,
                    )?;
                    output_row += 1;
                }
                "deleted" => {
                    let values = change.old_values.join(" | ");
                    write_row(
                        changes,
                        output_row,
                        &[
                            "ROW_DELETED",
                            &sheet.sheet,
                            &change.row,
                            "",
                            &values,
                            "",
                            "",
                        ],
                        &deleted_format,
                    )?;
                    output_row += 1;
                }
                "modified" => {
                    for cell in &change.cells {
                        write_row(
                            changes,
                            output_row,
                            &[
                                "CELL_MODIFIED",
                                &sheet.sheet,
                                &change.row,
                                &cell.column,
                                &cell.old,
                                &cell.new,
                                "",
                            ],
                            &modified_format,
                        )?;
                        output_row += 1;
                    }
                }
                _ => {}
            }
        }
    }
    changes.set_column_width(0, 18)?;
    changes.set_column_width(1, 18)?;
    changes.set_column_width(2, 18)?;
    changes.set_column_width(3, 24)?;
    changes.set_column_width(4, 42)?;
    changes.set_column_width(5, 42)?;
    changes.autofilter(0, 0, output_row.saturating_sub(1), 6)?;

    let issues = workbook.add_worksheet();
    issues.set_name("Issues")?;
    for (column, name) in [
        "Severity",
        "Code",
        "Sheet",
        "Row",
        "Column",
        "Message",
        "Reference",
    ]
    .into_iter()
    .enumerate()
    {
        issues.write_string_with_format(0, column as u16, name, &header_format)?;
    }
    for (index, issue) in result.issues.iter().enumerate() {
        let row = index as u32 + 1;
        issues.write_string(row, 0, &issue.severity)?;
        issues.write_string(row, 1, &issue.code)?;
        issues.write_string(row, 2, issue.sheet.as_deref().unwrap_or(""))?;
        issues.write_string(
            row,
            3,
            issue.row.map(|value| value.to_string()).unwrap_or_default(),
        )?;
        issues.write_string(row, 4, issue.column.as_deref().unwrap_or(""))?;
        issues.write_string(row, 5, &issue.message)?;
        issues.write_string(row, 6, issue.reference.as_deref().unwrap_or(""))?;
    }
    issues.set_column_width(0, 12)?;
    issues.set_column_width(1, 28)?;
    issues.set_column_width(2, 18)?;
    issues.set_column_width(3, 10)?;
    issues.set_column_width(4, 24)?;
    issues.set_column_width(5, 70)?;
    issues.set_column_width(6, 40)?;
    workbook.save(output_path)?;
    Ok(())
}

fn write_row(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    values: &[&str],
    format: &Format,
) -> Result<()> {
    for (column, value) in values.iter().enumerate() {
        worksheet.write_string_with_format(row, column as u16, *value, format)?;
    }
    Ok(())
}

fn format_for_kind<'a>(
    kind: &str,
    added: &'a Format,
    deleted: &'a Format,
    modified: &'a Format,
) -> &'a Format {
    match kind {
        "added" => added,
        "deleted" => deleted,
        _ => modified,
    }
}

pub fn write_merge_excel(
    result: &MergeResult,
    output_path: impl AsRef<Path>,
    allow_conflicts: bool,
) -> Result<()> {
    if result.has_unresolved_conflicts() && !allow_conflicts {
        result.require_resolved()?;
    }
    let mut workbook = XlsxWorkbook::new();
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#4472C4");
    let conflict_format = Format::new().set_background_color("#FFC7CE");
    let change_format = Format::new().set_background_color("#FFF2CC");

    for sheet in &result.workbook.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet.name)?;
        for (column, header) in sheet.headers.iter().enumerate() {
            for (row, value) in [&header.cn, &header.en, &header.type_name]
                .into_iter()
                .enumerate()
            {
                worksheet.write_string_with_format(
                    row as u32,
                    column as u16,
                    value,
                    &header_format,
                )?;
            }
            worksheet.set_column_width(column as u16, 18)?;
        }
        for (row_offset, row) in sheet.rows.iter().enumerate() {
            let output_row = (DATA_START_ROW - 1 + row_offset) as u32;
            for (column, cell) in row.cells.iter().enumerate() {
                if let Some(formula) = &cell.formula {
                    let formula = if formula.starts_with('=') {
                        formula.clone()
                    } else {
                        format!("={formula}")
                    };
                    worksheet.write_formula(output_row, column as u16, formula.as_str())?;
                } else if column == 0 && cell.value.is_empty() && !row.key.is_empty() {
                    write_typed_value(worksheet, output_row, 0, &row.key, sheet.headers.first())?;
                } else if !cell.value.is_empty() {
                    write_typed_value(
                        worksheet,
                        output_row,
                        column as u16,
                        &cell.value,
                        sheet.headers.get(column),
                    )?;
                }
            }
        }
        if !sheet.headers.is_empty() {
            worksheet.autofilter(
                2,
                0,
                DATA_START_ROW as u32 - 1 + sheet.rows.len() as u32,
                sheet.headers.len() as u16 - 1,
            )?;
        }
    }

    let review = workbook.add_worksheet();
    review.set_name("_MergeReview")?;
    for (row, (label, value)) in [
        ("Strategy", result.strategy.as_str().to_string()),
        ("Sheets", result.summary().sheets.to_string()),
        ("Rows", result.summary().rows.to_string()),
        ("Columns", result.summary().columns.to_string()),
        ("Changes", result.summary().changes.to_string()),
        ("Conflicts", result.summary().conflicts.to_string()),
    ]
    .into_iter()
    .enumerate()
    {
        review.write_string_with_format(row as u32, 0, label, &header_format)?;
        review.write_string(row as u32, 1, value)?;
    }
    review.set_column_width(0, 20)?;
    review.set_column_width(1, 80)?;

    let conflicts = workbook.add_worksheet();
    conflicts.set_name("_MergeConflicts")?;
    for (column, name) in [
        "Sheet", "Row", "Column", "Base", "Ours", "Theirs", "Result", "Note",
    ]
    .into_iter()
    .enumerate()
    {
        conflicts.write_string_with_format(0, column as u16, name, &header_format)?;
    }
    let mut conflict_row = 1u32;
    for change in result.conflicts() {
        write_row(
            conflicts,
            conflict_row,
            &[
                &change.sheet,
                &change.row,
                &change.column,
                &change.base,
                &change.ours,
                &change.theirs,
                &change.result,
                &change.note,
            ],
            &conflict_format,
        )?;
        conflict_row += 1;
    }
    conflicts.autofilter(0, 0, conflict_row.saturating_sub(1), 7)?;

    let changes = workbook.add_worksheet();
    changes.set_name("_MergeChanges")?;
    for (column, name) in [
        "Sheet", "Kind", "Row", "Column", "Base", "Ours", "Theirs", "Result", "Conflict", "Note",
    ]
    .into_iter()
    .enumerate()
    {
        changes.write_string_with_format(0, column as u16, name, &header_format)?;
    }
    for (index, change) in result.changes.iter().enumerate() {
        let row = index as u32 + 1;
        let format = if change.conflict {
            &conflict_format
        } else {
            &change_format
        };
        write_row(
            changes,
            row,
            &[
                &change.sheet,
                &change.kind,
                &change.row,
                &change.column,
                &change.base,
                &change.ours,
                &change.theirs,
                &change.result,
                if change.conflict { "true" } else { "false" },
                &change.note,
            ],
            format,
        )?;
    }
    changes.autofilter(0, 0, result.changes.len() as u32, 9)?;

    let issues = workbook.add_worksheet();
    issues.set_name("_MergeIssues")?;
    for (column, name) in [
        "Severity",
        "Code",
        "Sheet",
        "Row",
        "Column",
        "Message",
        "Reference",
    ]
    .into_iter()
    .enumerate()
    {
        issues.write_string_with_format(0, column as u16, name, &header_format)?;
    }
    for (index, issue) in result.issues.iter().enumerate() {
        let row = index as u32 + 1;
        write_row(
            issues,
            row,
            &[
                &issue.severity,
                &issue.code,
                issue.sheet.as_deref().unwrap_or(""),
                &issue.row.map(|value| value.to_string()).unwrap_or_default(),
                issue.column.as_deref().unwrap_or(""),
                &issue.message,
                issue.reference.as_deref().unwrap_or(""),
            ],
            &Format::default(),
        )?;
    }
    workbook.save(output_path)?;
    Ok(())
}

fn write_typed_value(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    column: u16,
    value: &str,
    header: Option<&crate::model::Header>,
) -> Result<()> {
    let type_name = header
        .map(|header| header.type_name.to_ascii_lowercase())
        .unwrap_or_default();
    if type_name.contains("int")
        && let Ok(number) = value.parse::<i64>()
    {
        worksheet.write_number(row, column, number as f64)?;
        return Ok(());
    }
    if (type_name.contains("float") || type_name.contains("double"))
        && let Ok(number) = value.parse::<f64>()
    {
        worksheet.write_number(row, column, number)?;
        return Ok(());
    }
    if type_name.contains("bool")
        && let Ok(boolean) = value.parse::<bool>()
    {
        worksheet.write_boolean(row, column, boolean)?;
        return Ok(());
    }
    worksheet.write_string(row, column, value)?;
    Ok(())
}
