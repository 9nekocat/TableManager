use std::path::{Path, PathBuf};
use std::time::Instant;

use table_core::diff_workbooks;
use table_core::model::read_workbook;

fn collect_xlsx(root: &Path, output: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().eq_ignore_ascii_case(".temp"))
            {
                continue;
            }
            collect_xlsx(&path, output);
        } else if path
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            && !path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
        {
            output.push(path);
        }
    }
}

fn main() {
    let root = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("A:/DataTable/Raw"));
    let mut files = Vec::new();
    collect_xlsx(&root, &mut files);
    files.sort();

    let total_started = Instant::now();
    let mut load_seconds = 0.0;
    let mut validate_seconds = 0.0;
    let mut diff_seconds = 0.0;
    let mut valid_files = 0usize;
    let mut errors = 0usize;
    for path in &files {
        let started = Instant::now();
        let mut workbook = match read_workbook(path) {
            Ok(workbook) => workbook,
            Err(_) => continue,
        };
        load_seconds += started.elapsed().as_secs_f64();
        valid_files += 1;

        let started = Instant::now();
        workbook.add_validation_issues();
        validate_seconds += started.elapsed().as_secs_f64();
        errors += workbook
            .issues
            .iter()
            .filter(|issue| issue.is_error())
            .count();

        let started = Instant::now();
        let cloned = workbook.clone();
        let _ = diff_workbooks(&workbook, &cloned);
        diff_seconds += started.elapsed().as_secs_f64();
    }
    println!(
        "{}",
        serde_json::json!({
            "root": root,
            "file_count": files.len(),
            "valid_files": valid_files,
            "errors": errors,
            "total_seconds": total_started.elapsed().as_secs_f64(),
            "load_seconds": load_seconds,
            "validate_seconds": validate_seconds,
            "diff_seconds": diff_seconds,
        })
    );
}
