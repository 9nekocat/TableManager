use std::path::PathBuf;
use std::time::Instant;

use table_core::diff_workbooks;
use table_core::format_text_diff;
use table_core::model::read_workbook;
use table_core::output::write_diff_excel;

fn main() {
    let mut args = std::env::args().skip(1);
    let path = PathBuf::from(args.next().expect("xlsx path"));
    let changes: usize = args
        .next()
        .and_then(|value| value.parse().ok())
        .unwrap_or(100);
    let output = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| path.with_file_name("rust_diff_bench.xlsx"));

    let started = Instant::now();
    let old = match read_workbook(&path) {
        Ok(workbook) => workbook,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(2);
        }
    };
    let load_seconds = started.elapsed().as_secs_f64();

    let started = Instant::now();
    let mut new = old.clone();
    let clone_seconds = started.elapsed().as_secs_f64();
    let mut changed = 0usize;
    'sheets: for sheet in &mut new.sheets {
        for row in &mut sheet.rows {
            for cell in row.cells.iter_mut().skip(1) {
                if changed >= changes {
                    break 'sheets;
                }
                cell.value = format!("benchmark-change-{changed}");
                cell.formula = None;
                changed += 1;
            }
        }
    }

    let started = Instant::now();
    let result = diff_workbooks(&old, &new);
    let diff_seconds = started.elapsed().as_secs_f64();
    let started = Instant::now();
    let text = format_text_diff(&result);
    let text_seconds = started.elapsed().as_secs_f64();
    let started = Instant::now();
    if let Err(error) = write_diff_excel(&result, &output) {
        eprintln!("{error}");
        std::process::exit(2);
    }
    let excel_seconds = started.elapsed().as_secs_f64();

    println!(
        "{}",
        serde_json::json!({
            "file": path,
            "changed_cells": changed,
            "load_seconds": load_seconds,
            "clone_seconds": clone_seconds,
            "diff_seconds": diff_seconds,
            "text_seconds": text_seconds,
            "excel_seconds": excel_seconds,
            "text_bytes": text.len(),
            "excel_output": output,
            "summary": result.summary,
            "issues": result.issues.len(),
        })
    );
}
