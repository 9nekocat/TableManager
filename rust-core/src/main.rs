use std::path::{Path, PathBuf};

use clap::{Parser, Subcommand, ValueEnum};
use serde::Serialize;
use serde_json::{Value, json};
use table_core::error::Result;
use table_core::model::read_workbook;
use table_core::output::write_diff_excel;
use table_core::{diff_files, format_text_diff};

#[derive(Parser)]
#[command(
    name = "table-core",
    version,
    about = "Rust xlsx text snapshot, validation, and diff core"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    Validate {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    Inspect {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    Diff {
        old: PathBuf,
        new: PathBuf,
        #[arg(long, value_enum, default_value = "text")]
        format: OutputFormat,
        #[arg(long)]
        output: Option<PathBuf>,
    },
}

#[derive(Clone, Copy, ValueEnum)]
enum OutputFormat {
    Text,
    Json,
    Excel,
}

#[derive(Serialize)]
struct ValidationSummary {
    ok: bool,
    path: String,
    sheets: usize,
    rows: usize,
    columns: usize,
    formulas: usize,
    errors: usize,
    warnings: usize,
    issues: Vec<table_core::Issue>,
}

fn run(cli: Cli) -> Result<i32> {
    match cli.command {
        Command::Validate {
            file,
            json: as_json,
        } => {
            let mut workbook = read_workbook(&file)?;
            workbook.add_validation_issues();
            let errors = workbook
                .issues
                .iter()
                .filter(|issue| issue.is_error())
                .count();
            let warnings = workbook.issues.len() - errors;
            let formulas = workbook.formula_count();
            let summary = ValidationSummary {
                ok: errors == 0,
                path: workbook.path,
                sheets: workbook.sheets.len(),
                rows: workbook.sheets.iter().map(|sheet| sheet.rows.len()).sum(),
                columns: workbook
                    .sheets
                    .iter()
                    .map(|sheet| sheet.headers.len())
                    .sum(),
                formulas,
                errors,
                warnings,
                issues: workbook.issues,
            };
            if as_json {
                println!("{}", serde_json::to_string_pretty(&summary)?);
            } else {
                println!("validation: {}", if summary.ok { "OK" } else { "FAILED" });
                println!(
                    "sheets={} rows={} columns={} formulas={} errors={} warnings={}",
                    summary.sheets,
                    summary.rows,
                    summary.columns,
                    summary.formulas,
                    summary.errors,
                    summary.warnings
                );
                for issue in &summary.issues {
                    println!("- [{}] {}: {}", issue.severity, issue.code, issue.message);
                }
            }
            Ok(if summary.ok { 0 } else { 1 })
        }
        Command::Inspect {
            file,
            json: as_json,
        } => {
            let workbook = read_workbook(&file)?;
            let value = json!({
                "file": workbook.path,
                "sheets": workbook.sheets.iter().map(|sheet| {
                    json!({
                        "name": sheet.name,
                        "rows": sheet.rows.len(),
                        "columns": sheet.headers,
                        "formulas": sheet.rows.iter().flat_map(|row| row.cells.iter()).filter(|cell| cell.formula.is_some()).count(),
                        "row_keys": sheet.rows.iter().take(20).map(|row| row.key.clone()).collect::<Vec<_>>(),
                    })
                }).collect::<Vec<_>>(),
            });
            if as_json {
                println!("{}", serde_json::to_string_pretty(&value)?);
            } else {
                print_inspect(&value);
            }
            Ok(0)
        }
        Command::Diff {
            old,
            new,
            format,
            output,
        } => {
            let result = diff_files(&old, &new)?;
            match format {
                OutputFormat::Text => println!("{}", format_text_diff(&result)),
                OutputFormat::Json => println!("{}", serde_json::to_string_pretty(&result)?),
                OutputFormat::Excel => {
                    let output = output.unwrap_or_else(|| default_diff_path(&new));
                    write_diff_excel(&result, &output)?;
                    println!("{}", output.display());
                }
            }
            Ok(if result.issues.iter().any(|issue| issue.is_error()) {
                1
            } else {
                0
            })
        }
    }
}

fn default_diff_path(path: &Path) -> PathBuf {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("workbook");
    path.with_file_name(format!("{stem}_diff.xlsx"))
}

fn print_inspect(value: &Value) {
    if let Some(sheets) = value.get("sheets").and_then(Value::as_array) {
        for sheet in sheets {
            println!(
                "{}: {} rows, {} columns",
                sheet.get("name").and_then(Value::as_str).unwrap_or(""),
                sheet.get("rows").and_then(Value::as_u64).unwrap_or(0),
                sheet
                    .get("columns")
                    .and_then(Value::as_array)
                    .map_or(0, Vec::len)
            );
        }
    }
}

fn main() {
    let cli = Cli::parse();
    match run(cli) {
        Ok(code) => std::process::exit(code),
        Err(error) => {
            eprintln!("error: {error}");
            std::process::exit(2);
        }
    }
}
