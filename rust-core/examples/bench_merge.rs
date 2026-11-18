use std::path::PathBuf;
use std::time::Instant;

use table_core::merge::{MergeStrategy, merge_workbooks};
use table_core::model::read_workbook;
use table_core::output::write_merge_excel;

fn mutate(workbook: &mut table_core::Workbook, start: usize, count: usize, prefix: &str) -> usize {
    let mut seen = 0usize;
    let mut changed = 0usize;
    'outer: for sheet in &mut workbook.sheets {
        for row in &mut sheet.rows {
            for cell in row.cells.iter_mut().skip(1) {
                if seen >= start && changed < count {
                    cell.value = format!("{prefix}-{changed}");
                    cell.formula = None;
                    changed += 1;
                }
                seen += 1;
                if changed >= count {
                    break 'outer;
                }
            }
        }
    }
    changed
}

fn main() {
    let mut args = std::env::args().skip(1);
    let path = PathBuf::from(args.next().expect("xlsx path"));
    let changes: usize = args
        .next()
        .and_then(|value| value.parse().ok())
        .unwrap_or(50);
    let output = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| path.with_file_name("rust_merge_bench.xlsx"));

    let started = Instant::now();
    let mut base = match read_workbook(&path) {
        Ok(workbook) => workbook,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(2);
        }
    };
    let load_seconds = started.elapsed().as_secs_f64();
    base.add_validation_issues();
    let mut ours = base.clone();
    let mut theirs = base.clone();
    let ours_changed = mutate(&mut ours, 0, changes, "ours");
    let theirs_changed = mutate(&mut theirs, changes, changes, "theirs");

    let started = Instant::now();
    let result = merge_workbooks(&base, &ours, &theirs, MergeStrategy::Ours);
    let merge_seconds = started.elapsed().as_secs_f64();
    let started = Instant::now();
    if let Err(error) = write_merge_excel(&result, &output, true) {
        eprintln!("{error}");
        std::process::exit(2);
    }
    let output_seconds = started.elapsed().as_secs_f64();
    println!(
        "{}",
        serde_json::json!({
            "file": path,
            "load_seconds": load_seconds,
            "merge_seconds": merge_seconds,
            "output_seconds": output_seconds,
            "total_after_load_seconds": merge_seconds + output_seconds,
            "ours_changed": ours_changed,
            "theirs_changed": theirs_changed,
            "summary": result.summary(),
            "output": output,
        })
    );
}
