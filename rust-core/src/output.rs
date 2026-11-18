use std::collections::{BTreeMap, HashSet};
use std::path::Path;

use rust_xlsxwriter::{Color, Format, Workbook as XlsxWorkbook};

use crate::diff::{DiffResult, RowChange, SheetDiff};
use crate::error::Result;
use crate::localization::DiffReportLabels;
use crate::merge::MergeResult;
use crate::model::DATA_START_ROW;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiffExcelLayout {
    Compact,
    Arranged,
}

pub fn write_diff_excel(result: &DiffResult, output_path: impl AsRef<Path>) -> Result<()> {
    write_diff_excel_with_layout(
        result,
        output_path,
        &DiffReportLabels::default(),
        DiffExcelLayout::Arranged,
    )
}

pub fn write_diff_excel_with_labels(
    result: &DiffResult,
    output_path: impl AsRef<Path>,
    labels: &DiffReportLabels,
) -> Result<()> {
    write_diff_excel_with_layout(result, output_path, labels, DiffExcelLayout::Arranged)
}

pub fn write_diff_excel_with_layout(
    result: &DiffResult,
    output_path: impl AsRef<Path>,
    labels: &DiffReportLabels,
    layout: DiffExcelLayout,
) -> Result<()> {
    if layout == DiffExcelLayout::Arranged {
        return write_arranged_diff_excel(result, output_path, labels);
    }
    let mut workbook = XlsxWorkbook::new();
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#4472C4");
    let title_format = Format::new()
        .set_bold()
        .set_font_size(16)
        .set_font_color("#1F3864");
    let section_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#5B9BD5");
    let summary_label_format = Format::new()
        .set_bold()
        .set_background_color("#D9EAF7")
        .set_text_wrap();
    let body_format = Format::new()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let added_format = Format::new()
        .set_background_color("#E2F0D9")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let deleted_format = Format::new()
        .set_background_color("#FCE4D6")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let old_format = Format::new()
        .set_background_color("#FFF2CC")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let new_format = Format::new()
        .set_background_color("#DDEBF7")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let old_changed_format = Format::new()
        .set_background_color("#FFF2CC")
        .set_font_color("#C00000")
        .set_bold()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let new_changed_format = Format::new()
        .set_background_color("#DDEBF7")
        .set_font_color("#006100")
        .set_bold()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let column_format = Format::new()
        .set_background_color("#EADCF8")
        .set_font_color("#5A2D82")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let issue_error_format = Format::new()
        .set_background_color("#FFC7CE")
        .set_text_wrap();

    let report = workbook.add_worksheet();
    report.set_name(safe_report_sheet_name(&labels.report_title))?;
    let mut output_row = write_diff_summary(
        report,
        0,
        result,
        labels,
        &title_format,
        &header_format,
        &summary_label_format,
        &body_format,
    )?;

    let changed_sheets = result
        .sheets
        .iter()
        .filter(|sheet| sheet.has_changes())
        .collect::<Vec<_>>();
    if changed_sheets.is_empty() {
        report.write_string_with_format(output_row, 0, &labels.no_differences, &section_format)?;
    } else {
        for sheet in changed_sheets {
            output_row = write_sheet_section(
                report,
                output_row,
                sheet,
                labels,
                &header_format,
                &section_format,
                &body_format,
                &added_format,
                &deleted_format,
                &old_format,
                &new_format,
                &old_changed_format,
                &new_changed_format,
                &column_format,
            )?;
        }
    }

    if !result.issues.is_empty() {
        output_row += 1;
        write_issue_section(
            report,
            output_row,
            result,
            labels,
            &header_format,
            &section_format,
            &body_format,
            &issue_error_format,
        )?;
    }

    report.set_column_width(0, 20)?;
    let max_columns = result
        .sheets
        .iter()
        .map(|sheet| sheet.column_alignments.len())
        .max()
        .unwrap_or(0);
    for column in 1..=max_columns {
        report.set_column_width(column as u16, 24)?;
    }
    workbook.save(output_path)?;
    Ok(())
}

fn write_arranged_diff_excel(
    result: &DiffResult,
    output_path: impl AsRef<Path>,
    labels: &DiffReportLabels,
) -> Result<()> {
    let mut workbook = XlsxWorkbook::new();
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#4472C4");
    let title_format = Format::new()
        .set_bold()
        .set_font_size(16)
        .set_font_color("#1F3864");
    let section_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color("#5B9BD5");
    let summary_label_format = Format::new()
        .set_bold()
        .set_background_color("#D9EAF7")
        .set_text_wrap();
    let body_format = Format::new()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let added_format = Format::new()
        .set_background_color("#E2F0D9")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let deleted_format = Format::new()
        .set_background_color("#FCE4D6")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let old_format = Format::new()
        .set_background_color("#FFF2CC")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let new_format = Format::new()
        .set_background_color("#DDEBF7")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let old_changed_format = Format::new()
        .set_background_color("#FFF2CC")
        .set_font_color("#C00000")
        .set_bold()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let new_changed_format = Format::new()
        .set_background_color("#DDEBF7")
        .set_font_color("#006100")
        .set_bold()
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let column_format = Format::new()
        .set_background_color("#EADCF8")
        .set_font_color("#5A2D82")
        .set_text_wrap()
        .set_align(rust_xlsxwriter::FormatAlign::Top);
    let issue_error_format = Format::new()
        .set_background_color("#FFC7CE")
        .set_text_wrap();

    let mut used_sheet_names = HashSet::new();
    let changed_sheets = result
        .sheets
        .iter()
        .filter(|sheet| sheet.has_changes())
        .collect::<Vec<_>>();
    for sheet in &changed_sheets {
        let sheet_name = unique_sheet_name(&sheet.sheet, &mut used_sheet_names);
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet_name)?;
        write_arranged_sheet(
            worksheet,
            sheet,
            labels,
            &header_format,
            &body_format,
            &added_format,
            &deleted_format,
            &old_format,
            &new_format,
            &old_changed_format,
            &new_changed_format,
            &column_format,
        )?;
    }

    let report_name = unique_sheet_name(&labels.report_title, &mut used_sheet_names);
    let report = workbook.add_worksheet();
    report.set_name(&report_name)?;
    let mut output_row = write_diff_summary(
        report,
        0,
        result,
        labels,
        &title_format,
        &header_format,
        &summary_label_format,
        &body_format,
    )?;
    if changed_sheets.is_empty() {
        report.write_string_with_format(output_row, 0, &labels.no_differences, &section_format)?;
        output_row += 2;
    }
    if !result.issues.is_empty() {
        write_arranged_issue_section(
            report,
            output_row,
            result,
            labels,
            &header_format,
            &section_format,
            &body_format,
            &issue_error_format,
        )?;
    }
    report.set_column_width(0, 14)?;
    report.set_column_width(1, 28)?;
    report.set_column_width(2, 18)?;
    report.set_column_width(3, 14)?;
    report.set_column_width(4, 24)?;
    report.set_column_width(5, 70)?;
    report.set_column_width(6, 30)?;
    workbook.save(output_path)?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_arranged_sheet(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    sheet: &SheetDiff,
    labels: &DiffReportLabels,
    header_format: &Format,
    body_format: &Format,
    added_format: &Format,
    deleted_format: &Format,
    old_format: &Format,
    new_format: &Format,
    old_changed_format: &Format,
    new_changed_format: &Format,
    column_format: &Format,
) -> Result<()> {
    let columns = &sheet.column_alignments;
    worksheet.write_string_with_format(0, 0, &labels.chinese_header, header_format)?;
    worksheet.write_string_with_format(1, 0, &labels.english_header, header_format)?;
    worksheet.write_string_with_format(2, 0, &labels.field_type, header_format)?;
    for (index, column) in columns.iter().enumerate() {
        let header = column
            .new_header
            .as_ref()
            .or(column.old_header.as_ref())
            .cloned()
            .unwrap_or_default();
        write_blank_or_string(worksheet, 0, index as u16 + 1, &header.cn, header_format)?;
        write_blank_or_string(worksheet, 1, index as u16 + 1, &header.en, header_format)?;
        write_blank_or_string(
            worksheet,
            2,
            index as u16 + 1,
            &header.type_name,
            header_format,
        )?;
    }
    worksheet.set_column_width(0, 18)?;
    for column in 0..columns.len() {
        worksheet.set_column_width(column as u16 + 1, 22)?;
    }
    worksheet.set_freeze_panes(4, 1)?;

    let mut row = 4u32;
    for change in &sheet.column_changes {
        row = write_column_change(worksheet, row, sheet, change, labels, column_format)?;
    }
    if !sheet.column_changes.is_empty() && !sheet.row_changes.is_empty() {
        row += 1;
    }
    for change in &sheet.row_changes {
        row = write_row_change(
            worksheet,
            row,
            sheet,
            change,
            labels,
            body_format,
            added_format,
            deleted_format,
            old_format,
            new_format,
            old_changed_format,
            new_changed_format,
        )?;
        row += 1;
    }
    Ok(())
}

struct ArrangedIssueRow {
    severity: String,
    code: String,
    sheet: String,
    row: String,
    column: String,
    message: String,
    reference: String,
}

#[allow(clippy::too_many_arguments)]
fn write_arranged_issue_section(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    result: &DiffResult,
    labels: &DiffReportLabels,
    header_format: &Format,
    section_format: &Format,
    body_format: &Format,
    error_format: &Format,
) -> Result<()> {
    worksheet.write_string_with_format(row, 0, &labels.issues, section_format)?;
    let mut output_row = row + 1;
    for (column, label) in [
        &labels.severity,
        &labels.code,
        &labels.sheet,
        &labels.row,
        &labels.column,
        &labels.message,
        &labels.reference,
    ]
    .into_iter()
    .enumerate()
    {
        worksheet.write_string_with_format(output_row, column as u16, label, header_format)?;
    }
    output_row += 1;

    for issue in arranged_issues(result, labels) {
        let format = if issue.severity == labels.error {
            error_format
        } else {
            body_format
        };
        let values = [
            issue.severity.as_str(),
            issue.code.as_str(),
            issue.sheet.as_str(),
            issue.row.as_str(),
            issue.column.as_str(),
            issue.message.as_str(),
            issue.reference.as_str(),
        ];
        for (column, value) in values.into_iter().enumerate() {
            write_blank_or_string(worksheet, output_row, column as u16, value, format)?;
        }
        output_row += 1;
    }
    Ok(())
}

fn arranged_issues(result: &DiffResult, labels: &DiffReportLabels) -> Vec<ArrangedIssueRow> {
    let mut blank_rows = BTreeMap::<String, Vec<usize>>::new();
    let mut issues = Vec::new();
    let mut seen = HashSet::new();
    for issue in &result.issues {
        if issue.code == "blank_row_removed"
            && let Some(row) = issue.row
        {
            blank_rows
                .entry(issue.sheet.clone().unwrap_or_default())
                .or_default()
                .push(row);
            continue;
        }
        let key = format!(
            "{}\u{1f}{}\u{1f}{}\u{1f}{}\u{1f}{}\u{1f}{}\u{1f}{}",
            issue.severity,
            issue.code,
            issue.sheet.as_deref().unwrap_or(""),
            issue.row.unwrap_or_default(),
            issue.column.as_deref().unwrap_or(""),
            issue.message,
            issue.reference.as_deref().unwrap_or("")
        );
        if !seen.insert(key) {
            continue;
        }
        let severity = match issue.severity.as_str() {
            "error" => labels.error.clone(),
            "warning" => labels.warning.clone(),
            _ => labels.info.clone(),
        };
        issues.push(ArrangedIssueRow {
            severity,
            code: issue.code.clone(),
            sheet: issue.sheet.clone().unwrap_or_default(),
            row: issue.row.map(|value| value.to_string()).unwrap_or_default(),
            column: issue.column.clone().unwrap_or_default(),
            message: issue.message.clone(),
            reference: issue.reference.clone().unwrap_or_default(),
        });
    }
    for (sheet, rows) in blank_rows {
        let count = rows.len();
        issues.push(ArrangedIssueRow {
            severity: labels.warning.clone(),
            code: "blank_row_removed".to_string(),
            sheet,
            row: format_row_ranges(rows),
            column: String::new(),
            message: labels
                .blank_rows_removed_group
                .replace("{count}", &count.to_string()),
            reference: String::new(),
        });
    }
    issues.sort_by_key(|issue| {
        let severity = if issue.severity == labels.error { 0 } else { 1 };
        (severity, issue.sheet.clone(), issue.code.clone())
    });
    issues
}

fn format_row_ranges(mut rows: Vec<usize>) -> String {
    rows.sort_unstable();
    rows.dedup();
    let mut ranges = Vec::new();
    let mut start = None;
    let mut end = 0usize;
    for row in rows {
        match start {
            None => {
                start = Some(row);
                end = row;
            }
            Some(_) if row == end + 1 => {
                end = row;
            }
            Some(value) => {
                ranges.push(if value == end {
                    value.to_string()
                } else {
                    format!("{value}-{end}")
                });
                start = Some(row);
                end = row;
            }
        }
    }
    if let Some(value) = start {
        ranges.push(if value == end {
            value.to_string()
        } else {
            format!("{value}-{end}")
        });
    }
    ranges.join(", ")
}

fn unique_sheet_name(requested: &str, used: &mut HashSet<String>) -> String {
    let base = safe_report_sheet_name(requested);
    let mut candidate = base.clone();
    let mut suffix = 2usize;
    while used.contains(&candidate.to_ascii_lowercase()) {
        let suffix_text = format!(" ({suffix})");
        let keep = 31usize.saturating_sub(suffix_text.chars().count());
        candidate = format!(
            "{}{}",
            base.chars().take(keep).collect::<String>(),
            suffix_text
        );
        suffix += 1;
    }
    used.insert(candidate.to_ascii_lowercase());
    candidate
}

#[allow(clippy::too_many_arguments)]
fn write_diff_summary(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    mut row: u32,
    result: &DiffResult,
    labels: &DiffReportLabels,
    title_format: &Format,
    header_format: &Format,
    label_format: &Format,
    body_format: &Format,
) -> Result<u32> {
    worksheet.write_string_with_format(row, 0, &labels.report_title, title_format)?;
    row += 2;
    worksheet.write_string_with_format(row, 0, &labels.summary, header_format)?;
    row += 1;
    let has_changes = if result.has_changes {
        labels.yes.as_str()
    } else {
        labels.no.as_str()
    };
    let entries = [
        (labels.old_file.as_str(), result.old.as_str()),
        (labels.new_file.as_str(), result.new.as_str()),
        (
            labels.comparison.as_str(),
            if result.comparison == "text_snapshot_with_cached_formula_values" {
                labels.comparison_cached_values.as_str()
            } else {
                result.comparison.as_str()
            },
        ),
        (labels.has_changes.as_str(), has_changes),
        (labels.rows_added.as_str(), ""),
        (labels.rows_deleted.as_str(), ""),
        (labels.rows_modified.as_str(), ""),
        (labels.columns_added.as_str(), ""),
        (labels.columns_deleted.as_str(), ""),
        (labels.columns_modified.as_str(), ""),
        (labels.columns_moved.as_str(), ""),
        (labels.cells_modified.as_str(), ""),
    ];
    let counts = [
        result.summary.rows_added.to_string(),
        result.summary.rows_deleted.to_string(),
        result.summary.rows_modified.to_string(),
        result.summary.columns_added.to_string(),
        result.summary.columns_deleted.to_string(),
        result.summary.columns_modified.to_string(),
        result.summary.columns_moved.to_string(),
        result.summary.cells_modified.to_string(),
    ];
    for (index, (label, value)) in entries.into_iter().enumerate() {
        worksheet.write_string_with_format(row, 0, label, label_format)?;
        let value = if index >= 4 {
            counts[index - 4].as_str()
        } else {
            value
        };
        write_blank_or_string(worksheet, row, 1, value, body_format)?;
        row += 1;
    }
    Ok(row + 1)
}

#[allow(clippy::too_many_arguments)]
fn write_sheet_section(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    mut row: u32,
    sheet: &SheetDiff,
    labels: &DiffReportLabels,
    header_format: &Format,
    section_format: &Format,
    body_format: &Format,
    added_format: &Format,
    deleted_format: &Format,
    old_format: &Format,
    new_format: &Format,
    old_changed_format: &Format,
    new_changed_format: &Format,
    column_format: &Format,
) -> Result<u32> {
    worksheet.write_string_with_format(
        row,
        0,
        format!("{}：{}", labels.worksheet, sheet.sheet),
        section_format,
    )?;
    row += 2;

    let columns = &sheet.column_alignments;
    worksheet.write_string_with_format(row, 0, &labels.chinese_header, header_format)?;
    for (index, column) in columns.iter().enumerate() {
        let header = column
            .new_header
            .as_ref()
            .or(column.old_header.as_ref())
            .cloned()
            .unwrap_or_default();
        write_blank_or_string(worksheet, row, index as u16 + 1, &header.cn, header_format)?;
    }
    row += 1;
    worksheet.write_string_with_format(row, 0, &labels.english_header, header_format)?;
    for (index, column) in columns.iter().enumerate() {
        let header = column
            .new_header
            .as_ref()
            .or(column.old_header.as_ref())
            .cloned()
            .unwrap_or_default();
        write_blank_or_string(worksheet, row, index as u16 + 1, &header.en, header_format)?;
    }
    row += 1;
    worksheet.write_string_with_format(row, 0, &labels.field_type, header_format)?;
    for (index, column) in columns.iter().enumerate() {
        let header = column
            .new_header
            .as_ref()
            .or(column.old_header.as_ref())
            .cloned()
            .unwrap_or_default();
        write_blank_or_string(
            worksheet,
            row,
            index as u16 + 1,
            &header.type_name,
            header_format,
        )?;
    }
    row += 2;

    for change in &sheet.column_changes {
        row = write_column_change(worksheet, row, sheet, change, labels, column_format)?;
    }
    if !sheet.column_changes.is_empty() && !sheet.row_changes.is_empty() {
        row += 1;
    }

    for change in &sheet.row_changes {
        row = write_row_change(
            worksheet,
            row,
            sheet,
            change,
            labels,
            body_format,
            added_format,
            deleted_format,
            old_format,
            new_format,
            old_changed_format,
            new_changed_format,
        )?;
        row += 1;
    }
    Ok(row + 1)
}

#[allow(clippy::too_many_arguments)]
fn write_row_change(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    sheet: &SheetDiff,
    change: &RowChange,
    labels: &DiffReportLabels,
    body_format: &Format,
    added_format: &Format,
    deleted_format: &Format,
    old_format: &Format,
    new_format: &Format,
    old_changed_format: &Format,
    new_changed_format: &Format,
) -> Result<u32> {
    let old_values = aligned_values(sheet, &change.old_values, true);
    let new_values = aligned_values(sheet, &change.new_values, false);
    match change.kind.as_str() {
        "added" => {
            write_labeled_values(
                worksheet,
                row,
                &format!("[{}]", labels.added_row),
                &new_values,
                added_format,
                None,
                None,
            )?;
            Ok(row + 1)
        }
        "deleted" => {
            write_labeled_values(
                worksheet,
                row,
                &format!("[{}]", labels.deleted_row),
                &old_values,
                deleted_format,
                None,
                None,
            )?;
            Ok(row + 1)
        }
        "modified" => {
            let changed = old_values
                .iter()
                .zip(new_values.iter())
                .map(|(old, new)| old != new)
                .collect::<Vec<_>>();
            write_labeled_values(
                worksheet,
                row,
                &format!("[{}]", labels.old_row),
                &old_values,
                old_format,
                Some(&changed),
                Some(old_changed_format),
            )?;
            write_labeled_values(
                worksheet,
                row + 1,
                &format!("[{}]", labels.new_row),
                &new_values,
                new_format,
                Some(&changed),
                Some(new_changed_format),
            )?;
            Ok(row + 2)
        }
        _ => {
            write_labeled_values(
                worksheet,
                row,
                &change.kind,
                &new_values,
                body_format,
                None,
                None,
            )?;
            Ok(row + 1)
        }
    }
}

fn write_column_change(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    sheet: &SheetDiff,
    change: &crate::diff::ColumnChange,
    labels: &DiffReportLabels,
    format: &Format,
) -> Result<u32> {
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
    let (label, message) = match change.kind.as_str() {
        "added" => (
            labels.added_column.as_str(),
            format!("{}：{}", labels.added_column, new),
        ),
        "deleted" => (
            labels.deleted_column.as_str(),
            format!("{}：{}", labels.deleted_column, old),
        ),
        "moved" => (
            labels.moved_column.as_str(),
            format!("{}：{}", labels.moved_column, new),
        ),
        _ => (
            labels.modified_column.as_str(),
            format!("{}：{} → {}", labels.modified_column, old, new),
        ),
    };
    let target = sheet
        .column_alignments
        .iter()
        .position(|column| {
            (change.new_index.is_some() && column.new_index == change.new_index)
                || (change.old_index.is_some() && column.old_index == change.old_index)
        })
        .unwrap_or(0);
    write_blank_or_string(worksheet, row, 0, &format!("[{label}]"), format)?;
    for column in 0..sheet.column_alignments.len() {
        write_blank_or_string(worksheet, row, column as u16 + 1, "", format)?;
    }
    write_blank_or_string(worksheet, row, target as u16 + 1, &message, format)?;
    Ok(row + 1)
}

fn write_labeled_values(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    label: &str,
    values: &[String],
    format: &Format,
    changed: Option<&[bool]>,
    changed_format: Option<&Format>,
) -> Result<()> {
    write_blank_or_string(worksheet, row, 0, label, format)?;
    for (index, value) in values.iter().enumerate() {
        let cell_format = changed
            .and_then(|flags| flags.get(index).copied())
            .filter(|changed| *changed)
            .and(changed_format)
            .unwrap_or(format);
        write_blank_or_string(worksheet, row, index as u16 + 1, value, cell_format)?;
    }
    let height = estimated_row_height(values);
    if height > 18.0 {
        worksheet.set_row_height(row, height)?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_issue_section(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    result: &DiffResult,
    labels: &DiffReportLabels,
    header_format: &Format,
    section_format: &Format,
    body_format: &Format,
    error_format: &Format,
) -> Result<()> {
    worksheet.write_string_with_format(row, 0, &labels.issues, section_format)?;
    let mut output_row = row + 1;
    for (column, label) in [
        &labels.severity,
        &labels.code,
        &labels.sheet,
        &labels.row,
        &labels.column,
        &labels.message,
        &labels.reference,
    ]
    .into_iter()
    .enumerate()
    {
        worksheet.write_string_with_format(output_row, column as u16, label, header_format)?;
    }
    output_row += 1;
    for issue in &result.issues {
        let format = if issue.severity == "error" {
            error_format
        } else {
            body_format
        };
        let severity = match issue.severity.as_str() {
            "error" => &labels.error,
            "warning" => &labels.warning,
            _ => &labels.info,
        };
        let row_text = issue.row.map(|value| value.to_string()).unwrap_or_default();
        let values = [
            severity.as_str(),
            issue.code.as_str(),
            issue.sheet.as_deref().unwrap_or(""),
            row_text.as_str(),
            issue.column.as_deref().unwrap_or(""),
            issue.message.as_str(),
            issue.reference.as_deref().unwrap_or(""),
        ];
        for (column, value) in values.into_iter().enumerate() {
            write_blank_or_string(worksheet, output_row, column as u16, value, format)?;
        }
        output_row += 1;
    }
    Ok(())
}

fn aligned_values(sheet: &SheetDiff, values: &[String], old_side: bool) -> Vec<String> {
    sheet
        .column_alignments
        .iter()
        .map(|column| {
            let index = if old_side {
                column.old_index
            } else {
                column.new_index
            };
            index
                .and_then(|index| values.get(index))
                .cloned()
                .unwrap_or_default()
        })
        .collect()
}

fn estimated_row_height(values: &[String]) -> f64 {
    let lines = values
        .iter()
        .map(|value| {
            value
                .lines()
                .map(|line| (line.chars().count() as f64 / 22.0).ceil() as usize)
                .sum::<usize>()
                .max(1)
        })
        .max()
        .unwrap_or(1);
    18.0_f64.max((lines as f64 * 15.0).min(90.0))
}

fn write_blank_or_string(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    column: u16,
    value: &str,
    format: &Format,
) -> Result<()> {
    if value.is_empty() {
        worksheet.write_blank(row, column, format)?;
    } else {
        worksheet.write_string_with_format(row, column, value, format)?;
    }
    Ok(())
}

fn safe_report_sheet_name(name: &str) -> String {
    let mut sanitized = name
        .chars()
        .map(|character| {
            if matches!(character, '[' | ']' | ':' | '*' | '?' | '/' | '\\') {
                '_'
            } else {
                character
            }
        })
        .take(31)
        .collect::<String>();
    sanitized = sanitized.trim_matches('\'').trim().to_string();
    if sanitized.is_empty() {
        "比对报告".to_string()
    } else {
        sanitized
    }
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
    let literal_formula_format = Format::new().set_quote_prefix().set_num_format("@");

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
                    write_typed_value(
                        worksheet,
                        output_row,
                        0,
                        &row.key,
                        sheet.headers.first(),
                        &literal_formula_format,
                    )?;
                } else if !cell.value.is_empty() {
                    write_typed_value(
                        worksheet,
                        output_row,
                        column as u16,
                        &cell.value,
                        sheet.headers.get(column),
                        &literal_formula_format,
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
    literal_formula_format: &Format,
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
    let value = value
        .strip_prefix("'=")
        .map_or_else(|| value.to_string(), |rest| format!("={rest}"));
    if value.starts_with('=') {
        worksheet.write_string_with_format(row, column, &value, literal_formula_format)?;
    } else {
        worksheet.write_string(row, column, &value)?;
    }
    Ok(())
}
