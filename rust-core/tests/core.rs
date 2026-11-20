use calamine::{Data, Reader, open_workbook_auto};
use rust_xlsxwriter::Workbook as XlsxWorkbook;
use table_core::diff_workbooks;
use table_core::formula::analyze_formula;
use table_core::merge::{
    ConflictDefault, MergeStrategy, merge_workbooks, merge_workbooks_with_conflict_default,
};
use table_core::merge_review::apply_merge_review;
use table_core::model::{Cell, Header, Row, Sheet, Workbook};
use table_core::output::{
    write_diff_excel_with_layout, write_merge_excel, write_merge_review, write_merged_workbook,
};
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
    let wrapper_config = directory.join("svn-wrapper.json");
    let merge_path = directory.join("merged.xlsx");
    let package = directory.join("Table Manager");
    let wrappers = package.join("svn");
    std::fs::create_dir_all(&wrappers).unwrap();
    let diff_script = wrappers.join("svn-diff.ps1");
    let merge_script = wrappers.join("svn-merge.ps1");
    std::fs::copy(
        concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-diff.ps1"),
        &diff_script,
    )
    .unwrap();
    std::fs::copy(
        concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-merge.ps1"),
        &merge_script,
    )
    .unwrap();
    std::fs::write(directory.join("blocked"), "").unwrap();
    std::fs::write(
        &wrapper_config,
        r#"{"diff_output_dir":"${SOURCE_DIR}\\blocked","fallback_diff_output_dir":"${SOURCE_DIR}\\.temp","remember_output_dir":true,"merge_conflict_default":"theirs"}"#,
    )
    .unwrap();
    write_svn_fixture(&base_path, "Alpha");
    write_svn_fixture(&mine_path, "Ours");
    write_svn_fixture(&theirs_path, "Theirs");

    let run_diff = |old: &std::path::Path, new: &std::path::Path| {
        std::process::Command::new("powershell.exe")
            .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
            .arg(&diff_script)
            .arg(old)
            .arg(new)
            .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
            .env("TM_SVN_CONFIG", &wrapper_config)
            .env("TM_SVN_HEADLESS", "1")
            .env("TM_SVN_NO_OPEN", "1")
            .status()
            .unwrap()
    };
    assert!(run_diff(&base_path, &mine_path).success());
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
    let remembered: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(&wrapper_config).unwrap()).unwrap();
    assert_eq!(
        remembered["diff_output_dir"].as_str(),
        Some(r"${SOURCE_DIR}\.temp")
    );
    assert_eq!(
        remembered["merge_conflict_default"].as_str(),
        Some("theirs")
    );

    let history = directory.join("tortoise-history");
    std::fs::create_dir(&history).unwrap();
    std::fs::write(history.join(".temp"), "").unwrap();
    let old_revision = history.join("Card.xlsx-rev1.svn000.tmp.xlsx");
    let new_revision = history.join("Card.xlsx-rev2.svn000.tmp.xlsx");
    std::fs::copy(&base_path, &old_revision).unwrap();
    std::fs::copy(&mine_path, &new_revision).unwrap();
    assert!(run_diff(&old_revision, &new_revision).success());
    let package_reports = package.join(".temp");
    let history_report = std::fs::read_dir(&package_reports)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .find(|path| {
            path.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("diff-Card.xlsx-rev2")
        })
        .expect("history diff report in package .temp");
    let report = open_workbook_auto(&history_report).unwrap();
    assert_eq!(report.sheet_names(), vec!["#Data", "比对报告"]);
    let after_history: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(&wrapper_config).unwrap()).unwrap();
    assert_eq!(after_history, remembered);

    std::fs::copy(&mine_path, &merge_path).unwrap();
    let merge = std::process::Command::new("powershell.exe")
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
        .arg(&merge_script)
        .arg(&merge_path)
        .arg(&theirs_path)
        .arg(&mine_path)
        .arg(&base_path)
        .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
        .env("TM_SVN_CONFIG", &wrapper_config)
        .env_remove("TM_SVN_MERGE_REVIEW")
        .env("TM_SVN_HEADLESS", "1")
        .env("TM_SVN_NO_OPEN", "1")
        .status()
        .unwrap();
    assert_eq!(merge.code(), Some(2));
    assert_eq!(
        std::fs::read(&merge_path).unwrap(),
        std::fs::read(&mine_path).unwrap()
    );
    let review_path = std::fs::read_dir(&package_reports)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .find(|path| {
            path.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("merge-merged-")
        })
        .expect("conflict review in package .temp");
    let review = open_workbook_auto(&review_path).unwrap();
    assert!(review.sheet_names().iter().any(|name| name == "_MergeMeta"));

    let non_conflict_merge_path = directory.join("non-conflict-merged.xlsx");
    let non_conflict_review_path = directory.join("non-conflict-review.xlsx");
    write_svn_fixture(&mine_path, "Alpha");
    let merge = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-merge.ps1"),
        ])
        .arg(&non_conflict_merge_path)
        .arg(&theirs_path)
        .arg(&mine_path)
        .arg(&base_path)
        .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
        .env("TM_SVN_CONFIG", &wrapper_config)
        .env("TM_SVN_MERGE_REVIEW", &non_conflict_review_path)
        .env("TM_SVN_HEADLESS", "1")
        .env("TM_SVN_NO_OPEN", "1")
        .status()
        .unwrap();
    assert_eq!(merge.code(), Some(0));
    assert!(non_conflict_merge_path.is_file());
    assert!(!non_conflict_review_path.is_file());
    let _ = std::fs::remove_dir_all(directory);
}

#[cfg(windows)]
#[test]
fn svn_headless_diff_uses_internal_diff_for_text_files() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-svn-text-diff-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let old_path = directory.join("old.txt");
    let new_path = directory.join("new.txt");
    std::fs::write(&old_path, "old value\n").unwrap();
    std::fs::write(&new_path, "new value\n").unwrap();

    let output = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-diff.ps1"),
        ])
        .arg(&old_path)
        .arg(&new_path)
        .env("TM_SVN_HEADLESS", "1")
        .output()
        .unwrap();

    assert_eq!(output.status.code(), Some(0));
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("-old value"));
    assert!(stdout.contains("+new value"));
    let _ = std::fs::remove_dir_all(directory);
}

#[cfg(windows)]
#[test]
fn svn_merge_failure_writes_a_tip_message() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-svn-failure-tip-{}-{}",
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
    let review_path = directory.join("review.xlsx");
    let wrapper_config = directory.join("svn-wrapper.json");
    let tip_log = directory.join("tip.log");
    let failing_core = directory.join("failing-table-core.cmd");
    write_svn_fixture(&base_path, "Base");
    write_svn_fixture(&mine_path, "Mine");
    write_svn_fixture(&theirs_path, "Theirs");
    std::fs::write(&wrapper_config, r#"{"merge_conflict_default":"theirs"}"#).unwrap();
    std::fs::write(&failing_core, "@exit 7\r\n").unwrap();

    let status = std::process::Command::new("powershell.exe")
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
        .env("TABLE_CORE_EXE", &failing_core)
        .env("TM_SVN_CONFIG", &wrapper_config)
        .env("TM_SVN_MERGE_REVIEW", &review_path)
        .env("TM_SVN_NO_OPEN", "1")
        .env("TM_SVN_TIP_LOG", &tip_log)
        .status()
        .unwrap();

    assert_eq!(status.code(), Some(2));
    let message = std::fs::read_to_string(&tip_log).unwrap();
    assert!(message.contains("Table merge failed"));
    assert!(message.contains("exit code 7"));
    let _ = std::fs::remove_dir_all(directory);
}

#[cfg(windows)]
#[test]
fn svn_merge_marks_a_real_working_copy_conflict_resolved() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-svn-resolve-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let repo = directory.join("repo");
    let wc1 = directory.join("wc1");
    let wc2 = directory.join("wc2");
    let wrapper_config = directory.join("svn-wrapper.json");
    std::fs::write(&wrapper_config, r#"{"merge_conflict_default":"theirs"}"#).unwrap();

    assert!(
        std::process::Command::new("svnadmin")
            .arg("create")
            .arg(&repo)
            .status()
            .unwrap()
            .success()
    );
    let repo_url = format!("file:///{}", repo.to_string_lossy().replace('\\', "/"));
    for working_copy in [&wc1, &wc2] {
        assert!(
            std::process::Command::new("svn")
                .arg("checkout")
                .arg(&repo_url)
                .arg(working_copy)
                .status()
                .unwrap()
                .success()
        );
    }

    let file1 = wc1.join("Card.xlsx");
    write_svn_fixture(&file1, "Base");
    assert!(
        std::process::Command::new("svn")
            .arg("add")
            .arg(&file1)
            .status()
            .unwrap()
            .success()
    );
    assert!(
        std::process::Command::new("svn")
            .arg("commit")
            .arg("-m")
            .arg("base")
            .arg(&wc1)
            .status()
            .unwrap()
            .success()
    );

    let file2 = wc2.join("Card.xlsx");
    assert!(
        std::process::Command::new("svn")
            .arg("update")
            .arg(&wc2)
            .status()
            .unwrap()
            .success()
    );
    write_svn_fixture(&file2, "Theirs");
    assert!(
        std::process::Command::new("svn")
            .arg("commit")
            .arg("-m")
            .arg("theirs")
            .arg(&wc2)
            .status()
            .unwrap()
            .success()
    );

    write_svn_fixture(&file1, "Mine");
    assert!(
        std::process::Command::new("svn")
            .arg("update")
            .arg("--accept")
            .arg("postpone")
            .arg(&wc1)
            .status()
            .unwrap()
            .success()
    );
    let conflicted = String::from_utf8(
        std::process::Command::new("svn")
            .arg("status")
            .arg(&file1)
            .output()
            .unwrap()
            .stdout,
    )
    .unwrap();
    assert!(
        conflicted.starts_with('C'),
        "expected C status: {conflicted}"
    );

    let mut conflict_files = std::fs::read_dir(&wc1)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("Card.xlsx.r"))
        })
        .collect::<Vec<_>>();
    conflict_files.sort();
    let base_path = conflict_files.first().expect("base conflict file");
    let theirs_path = conflict_files.last().expect("theirs conflict file");

    let status = std::process::Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            concat!(env!("CARGO_MANIFEST_DIR"), "\\svn\\svn-merge.ps1"),
        ])
        .arg(base_path)
        .arg(theirs_path)
        .arg(base_path)
        .arg(&file1)
        .arg(&file1)
        .env("TABLE_CORE_EXE", env!("CARGO_BIN_EXE_table-core"))
        .env("TM_SVN_CONFIG", &wrapper_config)
        .env("TM_SVN_HEADLESS", "1")
        .status()
        .unwrap();
    assert_eq!(status.code(), Some(0));

    let resolved_status = String::from_utf8(
        std::process::Command::new("svn")
            .arg("status")
            .arg(&file1)
            .output()
            .unwrap()
            .stdout,
    )
    .unwrap();
    assert!(
        !resolved_status.starts_with('C'),
        "still conflicted: {resolved_status}"
    );
    assert!(
        resolved_status.starts_with('M'),
        "expected M status: {resolved_status}"
    );
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

fn write_model_workbook(path: &std::path::Path, model: &Workbook) {
    let mut workbook = XlsxWorkbook::new();
    for sheet in &model.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet.name).unwrap();
        for (column, header) in sheet.headers.iter().enumerate() {
            worksheet
                .write_string(0, column as u16, &header.cn)
                .unwrap();
            worksheet
                .write_string(1, column as u16, &header.en)
                .unwrap();
            worksheet
                .write_string(2, column as u16, &header.type_name)
                .unwrap();
        }
        for (row_offset, row) in sheet.rows.iter().enumerate() {
            let excel_row = 3 + row_offset as u32;
            for (column, cell) in row.cells.iter().enumerate() {
                if let Some(formula) = &cell.formula {
                    let formula = if formula.starts_with('=') {
                        formula.clone()
                    } else {
                        format!("={formula}")
                    };
                    worksheet
                        .write_formula(excel_row, column as u16, formula.as_str())
                        .unwrap();
                } else if !cell.value.is_empty() {
                    worksheet
                        .write_string(excel_row, column as u16, &cell.value)
                        .unwrap();
                }
            }
        }
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
fn empty_conflict_default_does_not_delete_a_sheet_modified_by_theirs() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Base")])],
    );
    let ours = Workbook {
        path: "ours".to_string(),
        sheets: Vec::new(),
        issues: Vec::new(),
    };
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Remote")])],
    );

    let result = merge_workbooks_with_conflict_default(
        &base,
        &ours,
        &theirs,
        MergeStrategy::Fail,
        ConflictDefault::Empty,
    );

    assert!(
        result
            .conflicts()
            .any(|change| change.kind == "sheet_delete_conflict")
    );
    let sheet = result
        .workbook
        .sheet("#Data")
        .expect("empty should fall back to theirs for structural deletion conflicts");
    assert_eq!(sheet.rows[0].cells[1].value, "Remote");
}

#[test]
fn merge_review_surfaces_structural_conflicts_before_confirmation() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Base")])],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Local")])],
    );
    let theirs = Workbook {
        path: "theirs".to_string(),
        sheets: Vec::new(),
        issues: Vec::new(),
    };
    let result = merge_workbooks_with_conflict_default(
        &base,
        &ours,
        &theirs,
        MergeStrategy::Fail,
        ConflictDefault::Theirs,
    );
    assert!(result.preview_rows.is_empty());

    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-structural-review-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_merge_review(&result, &path).unwrap();
    let mut workbook = open_workbook_auto(&path).unwrap();
    let names = workbook.sheet_names();
    assert!(names.contains(&"合并冲突".to_string()));
    assert!(names.contains(&"结构冲突".to_string()));
    let structural = workbook.worksheet_range("结构冲突").unwrap();
    assert_eq!(cell_text(&structural, 1, 0), "sheet_delete_conflict");
    assert_eq!(cell_text(&structural, 1, 7), "deleted");
    let _ = std::fs::remove_file(path);
}

#[test]
fn merge_review_uses_three_way_rows_and_theirs_default() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Base")])],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Mine")])],
    );
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Remote")])],
    );
    let result = merge_workbooks_with_conflict_default(
        &base,
        &ours,
        &theirs,
        MergeStrategy::Fail,
        ConflictDefault::Theirs,
    );
    assert_eq!(result.preview_rows.len(), 1);
    assert_eq!(result.preview_rows[0].final_cells[1].value, "Remote");

    let path = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-merge-review-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    write_merge_review(&result, &path).unwrap();
    let mut workbook = open_workbook_auto(&path).unwrap();
    assert_eq!(workbook.sheet_names(), vec!["#Data", "_MergeMeta"]);
    let range = workbook.worksheet_range("#Data").unwrap();
    assert_eq!(cell_text(&range, 3, 0), "[基线]");
    assert_eq!(cell_text(&range, 4, 0), "[我的]");
    assert_eq!(cell_text(&range, 5, 0), "[仓库]");
    assert_eq!(cell_text(&range, 6, 0), "[最终]");
    assert_eq!(cell_text(&range, 6, 1), "[仓库]");
    assert_eq!(cell_text(&range, 6, 3), "Remote");
    let _ = std::fs::remove_file(path);
}

#[test]
fn merge_review_applies_default_and_configurable_side() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-merge-review-apply-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Base")])],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Mine")])],
    );
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Remote")])],
    );
    let base_path = directory.join("base.xlsx");
    let ours_path = directory.join("ours.xlsx");
    let theirs_path = directory.join("theirs.xlsx");
    write_model_workbook(&base_path, &base);
    write_model_workbook(&ours_path, &ours);
    write_model_workbook(&theirs_path, &theirs);

    for (default, expected) in [
        (ConflictDefault::Theirs, "Remote"),
        (ConflictDefault::Ours, "Mine"),
        (ConflictDefault::Empty, ""),
    ] {
        let result = merge_workbooks_with_conflict_default(
            &base,
            &ours,
            &theirs,
            MergeStrategy::Fail,
            default,
        );
        let review = directory.join(format!("review-{}.xlsx", default.as_str()));
        let output = directory.join(format!("resolved-{}.xlsx", default.as_str()));
        write_merge_review(&result, &review).unwrap();
        let applied = apply_merge_review(
            &review,
            &base_path,
            &ours_path,
            &theirs_path,
            MergeStrategy::Fail,
            default,
            &output,
        )
        .unwrap();
        assert_eq!(applied, 1);
        let resolved = table_core::model::read_workbook(&output).unwrap();
        assert_eq!(
            resolved.sheet("#Data").unwrap().rows[0].cells[1].value,
            expected
        );
    }
    let _ = std::fs::remove_dir_all(directory);
}

#[test]
fn merge_review_applies_delete_conflict_default() {
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-merge-review-delete-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha")]),
            ("B", &[("Row", "B"), ("Name", "Beta")]),
        ],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[
            ("A", &[("Row", "A"), ("Name", "Alpha")]),
            ("B", &[("Row", "B"), ("Name", "Beta local")]),
        ],
    );
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Alpha")])],
    );
    let base_path = directory.join("base.xlsx");
    let ours_path = directory.join("ours.xlsx");
    let theirs_path = directory.join("theirs.xlsx");
    write_model_workbook(&base_path, &base);
    write_model_workbook(&ours_path, &ours);
    write_model_workbook(&theirs_path, &theirs);

    for (default, expected_rows) in [
        (ConflictDefault::Theirs, 1usize),
        (ConflictDefault::Ours, 2usize),
    ] {
        let result = merge_workbooks_with_conflict_default(
            &base,
            &ours,
            &theirs,
            MergeStrategy::Fail,
            default,
        );
        assert_eq!(result.preview_rows[0].kind, "row_delete_conflict");
        let review = directory.join(format!("delete-review-{}.xlsx", default.as_str()));
        let output = directory.join(format!("delete-resolved-{}.xlsx", default.as_str()));
        write_merge_review(&result, &review).unwrap();
        apply_merge_review(
            &review,
            &base_path,
            &ours_path,
            &theirs_path,
            MergeStrategy::Fail,
            default,
            &output,
        )
        .unwrap();
        let resolved = table_core::model::read_workbook(&output).unwrap();
        assert_eq!(resolved.sheet("#Data").unwrap().rows.len(), expected_rows);
    }
    let _ = std::fs::remove_dir_all(directory);
}

#[test]
fn merged_workbook_and_review_export_share_data_sheet_values() {
    let base = table(
        "base",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Base")])],
    );
    let ours = table(
        "ours",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Local")])],
    );
    let theirs = table(
        "theirs",
        &[("行名", "Row", "FString"), ("名称", "Name", "FString")],
        &[("A", &[("Row", "A"), ("Name", "Local")])],
    );
    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Fail);
    let directory = std::env::current_dir()
        .unwrap()
        .join("target")
        .join(format!(
            "table-core-merge-writers-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
    std::fs::create_dir_all(&directory).unwrap();
    let clean_path = directory.join("clean.xlsx");
    let report_path = directory.join("report.xlsx");
    write_merged_workbook(&result, &clean_path).unwrap();
    write_merge_excel(&result, &report_path, true).unwrap();

    let mut clean = open_workbook_auto(&clean_path).unwrap();
    let mut report = open_workbook_auto(&report_path).unwrap();
    assert_eq!(clean.sheet_names(), vec!["#Data"]);
    assert!(report.sheet_names().contains(&"_MergeReview".to_string()));
    let clean_data = clean.worksheet_range("#Data").unwrap();
    let report_data = report.worksheet_range("#Data").unwrap();
    assert_eq!(cell_text(&clean_data, 3, 1), cell_text(&report_data, 3, 1));
    assert_eq!(cell_text(&clean_data, 3, 2), cell_text(&report_data, 3, 2));
    let _ = std::fs::remove_dir_all(directory);
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
