use table_core::diff_workbooks;
use table_core::formula::analyze_formula;
use table_core::model::{Cell, Header, Row, Sheet, Workbook};
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
