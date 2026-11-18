use calamine::{Data, Reader, open_workbook_auto};
use rust_xlsxwriter::Workbook as XlsxWorkbook;
use table_core::diff_workbooks;
use table_core::formula::analyze_formula;
use table_core::merge::{MergeStrategy, merge_workbooks};
use table_core::model::{Cell, Header, Row, Sheet, Workbook};
use table_core::output::{write_diff_excel_with_layout, write_merge_excel};
use table_core::snapshot::workbook_to_text;
use table_core::{DiffExcelLayout, DiffReportLabels, Issue};

fn header(cn: &str, en: &str) -> Header {
    Header::new(cn, en, "string")
}

fn workbook(cells: Vec<Vec<Cell>>) -> Workbook {
    Workbook {
        path: "memory.xlsx".to_string(),
        sheets: vec![Sheet {
            name: "#Data".to_string(),
            headers: vec![header("行名", "Row"), header("名称", "Name")],
            rows: vec![
                Row {
                    key: "A".to_string(),
                    occurrence: 1,
                    cells: cells[0].clone(),
                    source_row: Some(4),
                },
                Row {
                    key: "B".to_string(),
                    occurrence: 1,
                    cells: cells[1].clone(),
                    source_row: Some(5),
                },
            ],
        }],
        issues: Vec::new(),
    }
}

fn table(path: &str, headers: &[(&str, &str, &str)], rows: &[(&str, &[(&str, &str)])]) -> Workbook {
    let headers = headers
        .iter()
        .map(|(cn, en, type_name)| Header::new(*cn, *en, *type_name))
        .collect::<Vec<_>>();
    let rows = rows
        .iter()
        .enumerate()
        .map(|(row_index, (key, values))| {
            let mut cells = vec![Cell::default(); headers.len()];
            for (column, value) in *values {
                if let Some(index) = headers.iter().position(|header| header.en == *column) {
                    cells[index] = if value.starts_with('=') {
                        Cell {
                            value: String::new(),
                            formula: Some(value.trim_start_matches('=').to_string()),
                        }
                    } else {
                        Cell {
                            value: value.to_string(),
                            formula: None,
                        }
                    };
                }
            }
            Row {
                key: key.to_string(),
                occurrence: 1,
                cells,
                source_row: Some(4 + row_index),
            }
        })
        .collect();
    Workbook {
        path: path.to_string(),
        sheets: vec![Sheet {
            name: "#Data".to_string(),
            headers,
            rows,
        }],
        issues: Vec::new(),
    }
}

#[test]
fn diff_uses_text_snapshot_values() {
    let old = workbook(vec![
        vec![Cell::default(), Cell::default()],
        vec![Cell::default(), Cell::default()],
    ]);
    let new = workbook(vec![
        vec![Cell::default(), Cell::default()],
        vec![Cell::default(), Cell::default()],
    ]);
    let result = diff_workbooks(&old, &new);
    assert!(!result.has_changes);
}

#[test]
fn diff_excel_keeps_summary_and_original_cell_layout_on_one_page() {
    let old = table(
        "old.xlsx",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha"), ("Power", "1")]),
            ("B", &[("Row", "B"), ("Name", "Beta"), ("Power", "2")]),
            ("C", &[("Row", "C"), ("Name", "Gamma"), ("Power", "3")]),
        ],
    );
    let new = table(
        "new.xlsx",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha2"), ("Power", "1")]),
            ("B", &[("Row", "B"), ("Name", "Beta"), ("Power", "2")]),
            ("D", &[("Row", "D"), ("Name", "Delta"), ("Power", "4")]),
        ],
    );
    let result = diff_workbooks(&old, &new);
    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-diff-layout-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_diff_excel_with_layout(
        &result,
        &path,
        &DiffReportLabels::default(),
        DiffExcelLayout::Compact,
    )
    .unwrap();

    let mut workbook = open_workbook_auto(&path).unwrap();
    assert_eq!(workbook.sheet_names(), vec!["比对报告"]);
    let range = workbook.worksheet_range("比对报告").unwrap();
    assert!(range_contains(&range, "摘要"));
    assert!(range_contains(&range, "旧文件"));
    assert!(!range_contains(&range, "Alpha | Alpha2"));

    let old_row = find_row(&range, "[旧版]").expect("old version row");
    assert_eq!(cell_text(&range, old_row, 1), "A");
    assert_eq!(cell_text(&range, old_row, 2), "Alpha");
    assert_eq!(cell_text(&range, old_row, 3), "1");

    let new_row = find_row(&range, "[当前]").expect("current version row");
    assert_eq!(cell_text(&range, new_row, 1), "A");
    assert_eq!(cell_text(&range, new_row, 2), "Alpha2");
    assert_eq!(cell_text(&range, new_row, 3), "1");

    let added_row = find_row(&range, "[新增]").expect("added row");
    assert_eq!(cell_text(&range, added_row, 1), "D");
    assert_eq!(cell_text(&range, added_row, 2), "Delta");
    assert_eq!(cell_text(&range, added_row, 3), "4");

    let deleted_row = find_row(&range, "[删除]").expect("deleted row");
    assert_eq!(cell_text(&range, deleted_row, 1), "C");
    assert_eq!(cell_text(&range, deleted_row, 2), "Gamma");
    assert_eq!(cell_text(&range, deleted_row, 3), "3");

    let _ = std::fs::remove_file(path);
}

#[test]
fn arranged_diff_excel_uses_sheet_tabs_and_groups_blank_row_issues() {
    let old = table(
        "old.xlsx",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Alpha")])],
    );
    let new = table(
        "new.xlsx",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha")]),
            ("B", &[("Row", "B"), ("Name", "Beta")]),
        ],
    );
    let mut result = diff_workbooks(&old, &new);
    for row in [410, 411, 412, 500, 501] {
        result.issues.push(Issue::new(
            "blank_row_removed",
            "warning",
            "检测到空数据行，加载时已从模型中移除。",
            Some("#Data"),
            Some(row),
            None,
            None,
        ));
    }
    let important = Issue::new(
        "cross_row_ref",
        "error",
        "公式引用了其他数据行。",
        Some("#Data"),
        Some(189),
        Some("MechaId (机甲ID)"),
        Some("B187".to_string()),
    );
    result.issues.push(important.clone());
    result.issues.push(important);

    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-arranged-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_diff_excel_with_layout(
        &result,
        &path,
        &DiffReportLabels::default(),
        DiffExcelLayout::Arranged,
    )
    .unwrap();
    let mut workbook = open_workbook_auto(&path).unwrap();
    let names = workbook.sheet_names();
    assert_eq!(names, vec!["#Data", "比对报告"]);
    let data = workbook.worksheet_range("#Data").unwrap();
    assert_eq!(cell_text(&data, 0, 0), "中文字段名");
    assert_eq!(cell_text(&data, 1, 0), "英文字段名");
    assert_eq!(cell_text(&data, 2, 0), "字段类型");
    assert_eq!(cell_text(&data, 4, 0), "[新增]");
    assert_eq!(cell_text(&data, 4, 1), "B");

    let report = workbook.worksheet_range("比对报告").unwrap();
    let blank_rows = (0..report.height() as u32)
        .filter(|row| cell_text(&report, *row, 1) == "blank_row_removed")
        .collect::<Vec<_>>();
    assert_eq!(blank_rows.len(), 1);
    assert_eq!(cell_text(&report, blank_rows[0], 3), "410-412, 500-501");
    assert!(cell_text(&report, blank_rows[0], 5).contains("5 个空数据行"));
    let important_rows = (0..report.height() as u32)
        .filter(|row| cell_text(&report, *row, 1) == "cross_row_ref")
        .collect::<Vec<_>>();
    assert_eq!(important_rows.len(), 1);
    assert_eq!(cell_text(&report, important_rows[0], 0), "错误");

    let _ = std::fs::remove_file(path);
}

#[test]
fn diff_report_labels_support_partial_localization_overrides() {
    let labels: DiffReportLabels =
        serde_json::from_str(r#"{"report_title":"Custom diff"}"#).unwrap();
    assert_eq!(labels.report_title, "Custom diff");
    assert_eq!(labels.summary, "摘要");
    assert_eq!(labels.old_row, "旧版");
}

#[cfg(windows)]
#[test]
fn svn_headless_wrappers_generate_diff_and_merge_files() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-svn-headless-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let base_path = directory.join("base.xlsx");
    let mine_path = directory.join("mine.xlsx");
    let theirs_path = directory.join("theirs.xlsx");
    let merge_path = directory.join("merged.xlsx");
    write_svn_fixture(&base_path, "Alpha");
    write_svn_fixture(&mine_path, "Ours");
    write_svn_fixture(&theirs_path, "Theirs");

    let diff = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-diff.ps1"),
        ])
        .arg(&base_path)
        .arg(&mine_path)
        .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
        .env("TM_SVN_HEADLESS", "1")
        .env("TM_SVN_NO_OPEN", "1")
        .status()
        .unwrap();
    assert!(diff.success());
    let generated_diff = std::fs::read_dir(directory.join(".temp"))
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .find(|path| {
            path.extension()
                .is_some_and(|extension| extension == "xlsx")
        })
        .expect("default diff output next to source");
    assert!(generated_diff.is_file());

    let merge = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-merge.ps1"),
        ])
        .arg(&merge_path)
        .arg(&theirs_path)
        .arg(&mine_path)
        .arg(&base_path)
        .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
        .env("TM_SVN_HEADLESS", "1")
        .env("TM_SVN_NO_OPEN", "1")
        .status()
        .unwrap();
    assert_eq!(merge.code(), Some(2));
    assert!(merge_path.is_file());
    let _ = std::fs::remove_dir_all(directory);
}

#[cfg(windows)]
fn write_svn_fixture(path: &std::path::Path, name: &str) {
    let mut workbook = XlsxWorkbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("#Data").unwrap();
    for (column, value) in ["行名", "名称", "数值"].into_iter().enumerate() {
        worksheet.write_string(0, column as u16, value).unwrap();
    }
    for (column, value) in ["Row", "Name", "Power"].into_iter().enumerate() {
        worksheet.write_string(1, column as u16, value).unwrap();
    }
    for (column, value) in ["FString", "FString", "int32"].into_iter().enumerate() {
        worksheet.write_string(2, column as u16, value).unwrap();
    }
    for (column, value) in ["A", name, "1"].into_iter().enumerate() {
        worksheet.write_string(3, column as u16, value).unwrap();
    }
    workbook.save(path).unwrap();
}

fn range_contains(range: &calamine::Range<Data>, expected: &str) -> bool {
    range
        .cells()
        .any(|(_, _, value)| cell_value(value) == expected)
}

fn find_row(range: &calamine::Range<Data>, first_cell: &str) -> Option<u32> {
    (0..range.height() as u32).find(|row| cell_text(range, *row, 0) == first_cell)
}

fn cell_text(range: &calamine::Range<Data>, row: u32, column: u32) -> String {
    range
        .get_value((row, column))
        .map(cell_value)
        .unwrap_or_default()
}

fn cell_value(value: &Data) -> String {
    match value {
        Data::Empty => String::new(),
        Data::String(value) => value.clone(),
        Data::Float(value) if value.fract() == 0.0 => format!("{}", *value as i64),
        other => other.to_string(),
    }
}

#[test]
fn formula_without_cached_value_is_empty_and_warned() {
    let workbook = workbook(vec![
        vec![
            Cell::default(),
            Cell {
                value: String::new(),
                formula: Some("B4".to_string()),
            },
        ],
        vec![Cell::default(), Cell::default()],
    ]);
    let snapshot = workbook_to_text(&workbook);
    assert_eq!(snapshot.sheets[0].rows[0].cells[1], "");
    assert!(
        snapshot
            .issues
            .iter()
            .any(|issue| issue.code == "formula_value_unavailable")
    );
}

#[test]
fn formula_risk_distinguishes_safe_and_cross_row_references() {
    assert!(analyze_formula("C4&D4", 4).is_empty());
    let fixed_sheet = analyze_formula("'#Lookup'!$A$1", 4);
    assert!(
        fixed_sheet
            .iter()
            .all(|finding| finding.severity != "error"),
        "{fixed_sheet:?}"
    );
    assert!(
        analyze_formula("C5", 4)
            .iter()
            .any(|finding| finding.code == "cross_row_ref")
    );
    assert!(
        analyze_formula("'[1]#Main'!$C$3882", 4)
            .iter()
            .any(|finding| finding.code == "external_workbook_ref")
    );
}

#[test]
fn duplicate_row_keys_are_reported() {
    let mut workbook = workbook(vec![
        vec![Cell::default(), Cell::default()],
        vec![Cell::default(), Cell::default()],
    ]);
    workbook.sheets[0].rows[1].key = "A".to_string();
    workbook.add_validation_issues();
    assert!(
        workbook
            .issues
            .iter()
            .any(|issue| issue.code == "duplicate_row_key")
    );
}

#[test]
fn merge_combines_cells_rows_and_columns() {
    let base = table(
        "base",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha"), ("Power", "1")]),
            ("B", &[("Row", "B"), ("Name", "Beta"), ("Power", "2")]),
        ],
    );
    let ours = table(
        "ours",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
            ("辅助", "Helper", "FString"),
        ],
        &[
            (
                "A",
                &[
                    ("Row", "A"),
                    ("Name", "Alpha"),
                    ("Power", "10"),
                    ("Helper", "ours"),
                ],
            ),
            (
                "B",
                &[
                    ("Row", "B"),
                    ("Name", "Beta"),
                    ("Power", "2"),
                    ("Helper", ""),
                ],
            ),
            (
                "C",
                &[
                    ("Row", "C"),
                    ("Name", "Gamma"),
                    ("Power", "3"),
                    ("Helper", "new"),
                ],
            ),
        ],
    );
    let theirs = table(
        "theirs",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("Z", &[("Row", "Z"), ("Name", "Before"), ("Power", "0")]),
            ("A", &[("Row", "A"), ("Name", "Alpha2"), ("Power", "1")]),
            ("B", &[("Row", "B"), ("Name", "Beta"), ("Power", "2")]),
            ("D", &[("Row", "D"), ("Name", "Delta"), ("Power", "4")]),
        ],
    );

    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    assert!(result.ok());
    let sheet = result.workbook.sheet("#Data").unwrap();
    assert_eq!(
        sheet
            .headers
            .iter()
            .map(|header| header.en.as_str())
            .collect::<Vec<_>>(),
        vec!["Row", "Name", "Power", "Helper"]
    );
    assert_eq!(
        sheet
            .rows
            .iter()
            .map(|row| row.key.as_str())
            .collect::<Vec<_>>(),
        vec!["Z", "A", "B", "D", "C"]
    );
    let row_a = sheet.rows.iter().find(|row| row.key == "A").unwrap();
    assert_eq!(row_a.cells[1].value, "Alpha2");
    assert_eq!(row_a.cells[2].value, "10");
    assert_eq!(row_a.cells[3].value, "ours");
}

#[test]
fn merge_detects_and_resolves_cell_conflicts() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("数值", "Power", "int32")],
        &[("A", &[("Row", "A"), ("Power", "1")])],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("数值", "Power", "int32")],
        &[("A", &[("Row", "A"), ("Power", "2")])],
    );
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("数值", "Power", "int32")],
        &[("A", &[("Row", "A"), ("Power", "3")])],
    );

    let failed = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    assert!(failed.has_unresolved_conflicts());
    assert!(failed.require_resolved().is_err());

    let ours_result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Ours);
    assert!(ours_result.ok());
    assert_eq!(
        ours_result.workbook.sheet("#Data").unwrap().rows[0].cells[1].value,
        "2"
    );
    let theirs_result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Theirs);
    assert_eq!(
        theirs_result.workbook.sheet("#Data").unwrap().rows[0].cells[1].value,
        "3"
    );
}

#[test]
fn merge_respects_row_and_column_deletion_strategy() {
    let base = table(
        "base",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha"), ("Power", "1")]),
            ("B", &[("Row", "B"), ("Name", "Beta"), ("Power", "2")]),
        ],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Alpha")])],
    );
    let theirs = table(
        "theirs",
        &[
            ("行名", "Row", "FString"),
            ("名称", "Name", "FString"),
            ("数值", "Power", "int32"),
        ],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha"), ("Power", "1")]),
            (
                "B",
                &[("Row", "B"), ("Name", "Beta changed"), ("Power", "20")],
            ),
        ],
    );

    let failed = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    assert!(failed.has_unresolved_conflicts());
    assert_eq!(
        failed
            .workbook
            .sheet("#Data")
            .unwrap()
            .headers
            .iter()
            .map(|header| header.en.as_str())
            .collect::<Vec<_>>(),
        vec!["Row", "Name", "Power"]
    );
    assert_eq!(failed.workbook.sheet("#Data").unwrap().rows.len(), 2);

    let ours_result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Ours);
    assert!(ours_result.ok());
    assert_eq!(
        ours_result
            .workbook
            .sheet("#Data")
            .unwrap()
            .headers
            .iter()
            .map(|header| header.en.as_str())
            .collect::<Vec<_>>(),
        vec!["Row", "Name"]
    );
    assert_eq!(ours_result.workbook.sheet("#Data").unwrap().rows.len(), 1);

    let theirs_result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Theirs);
    assert!(theirs_result.ok());
    assert_eq!(theirs_result.workbook.sheet("#Data").unwrap().rows.len(), 2);
}

#[test]
fn merge_rebinds_row_local_formula_after_row_insert() {
    let base = table(
        "base",
        &[
            ("行名", "Row", "FString"),
            ("A", "A", "FString"),
            ("B", "B", "FString"),
            ("拼接", "Combined", "FString"),
        ],
        &[
            (
                "A",
                &[("Row", "A"), ("A", "x"), ("B", "y"), ("Combined", "=B4&C4")],
            ),
            (
                "B",
                &[("Row", "B"), ("A", "p"), ("B", "q"), ("Combined", "=B5&C5")],
            ),
        ],
    );
    let ours = base.clone();
    let theirs = table(
        "theirs",
        &[
            ("行名", "Row", "FString"),
            ("A", "A", "FString"),
            ("B", "B", "FString"),
            ("拼接", "Combined", "FString"),
        ],
        &[
            (
                "Z",
                &[("Row", "Z"), ("A", "z"), ("B", "w"), ("Combined", "=B4&C4")],
            ),
            (
                "A",
                &[("Row", "A"), ("A", "x"), ("B", "y"), ("Combined", "=B5&C5")],
            ),
            (
                "B",
                &[("Row", "B"), ("A", "p"), ("B", "q"), ("Combined", "=B6&C6")],
            ),
        ],
    );

    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    let sheet = result.workbook.sheet("#Data").unwrap();
    assert_eq!(sheet.rows[1].cells[3].formula.as_deref(), Some("B5&C5"));
    assert_eq!(sheet.rows[2].cells[3].formula.as_deref(), Some("B6&C6"));
}

#[test]
fn merge_excel_cleans_blank_rows_and_is_readable() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Alpha")])],
    );
    let ours = base.clone();
    let theirs = base.clone();
    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-merge-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_merge_excel(&result, &path, false).unwrap();
    assert!(path.exists());
    let reopened = table_core::model::read_workbook(&path).unwrap();
    assert_eq!(reopened.sheets[0].rows.len(), 1);
    let _ = std::fs::remove_file(path);
}

#[test]
fn merge_writes_literal_formula_source_as_quoted_text() {
    let mut base = table(
        "base",
        &[
            ("行名", "Row", "FString"),
            ("公式源", "FormulaSource", "FString"),
        ],
        &[("A", &[("Row", "A"), ("FormulaSource", "placeholder")])],
    );
    base.sheets[0].rows[0].cells[1] = Cell {
        value: "=A,B,C".to_string(),
        formula: None,
    };
    let ours = base.clone();
    let theirs = base.clone();
    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-literal-formula-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_merge_excel(&result, &path, false).unwrap();
    let reopened = table_core::model::read_workbook(&path).unwrap();
    let cell = &reopened.sheets[0].rows[0].cells[1];
    assert_eq!(cell.value, "=A,B,C");
    assert!(cell.formula.is_none());
    let _ = std::fs::remove_file(path);
}
