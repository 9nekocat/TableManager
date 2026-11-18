use table_core::diff_workbooks;
use table_core::formula::analyze_formula;
use table_core::merge::{MergeStrategy, merge_workbooks};
use table_core::model::{Cell, Header, Row, Sheet, Workbook};
use table_core::output::write_merge_excel;
use table_core::snapshot::workbook_to_text;

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
