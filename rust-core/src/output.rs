use std::path::Path;

use rust_xlsxwriter::{Color, Format, Workbook as XlsxWorkbook};

use crate::diff::DiffResult;
use crate::error::Result;

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
