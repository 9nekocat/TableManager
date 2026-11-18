use std::path::{Path, PathBuf};

use clap::{Parser, Subcommand, ValueEnum};
use serde::Serialize;
use serde_json::{Value, json};
use table_core::error::Result;
use table_core::localization::DiffReportLabels;
use table_core::merge::{MergeStrategy, format_merge_text, merge_files};
use table_core::model::read_workbook;
use table_core::output::{DiffExcelLayout, write_diff_excel_with_layout, write_merge_excel};
use table_core::svn_integration::install_svn_integration;
use table_core::{diff_files, format_text_diff};

#[derive(Parser)]
#[command(
    name = "table-core",
    version,
    about = "Excel table diff and three-way merge"
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
        /// Optional JSON file overriding localized Excel report labels.
        #[arg(long)]
        labels: Option<PathBuf>,
        /// Excel report layout: arranged per-sheet (default) or compact single-page.
        #[arg(long, value_enum, default_value = "arranged")]
        layout: ExcelLayoutArg,
    },
    Merge {
        base: PathBuf,
        ours: PathBuf,
        theirs: PathBuf,
        #[arg(long, value_enum, default_value = "fail")]
        strategy: StrategyArg,
        #[arg(long)]
        output: Option<PathBuf>,
        #[arg(long, value_enum, default_value = "text")]
        report: ReportFormat,
        #[arg(long)]
        allow_conflicts: bool,
    },
    /// Register this package as the Subversion diff and merge helper.
    InstallSvn {
        /// Override the Subversion user config file.
        #[arg(long)]
        config: Option<PathBuf>,
        #[arg(long)]
        dry_run: bool,
    },
}

#[derive(Clone, Copy, ValueEnum)]
enum OutputFormat {
    Text,
    Json,
    Excel,
}

#[derive(Clone, Copy, ValueEnum)]
enum ReportFormat {
    Text,
    Json,
}

#[derive(Clone, Copy, ValueEnum)]
enum ExcelLayoutArg {
    Compact,
    Arranged,
}

impl From<ExcelLayoutArg> for DiffExcelLayout {
    fn from(value: ExcelLayoutArg) -> Self {
        match value {
            ExcelLayoutArg::Compact => DiffExcelLayout::Compact,
            ExcelLayoutArg::Arranged => DiffExcelLayout::Arranged,
        }
    }
}

#[derive(Clone, Copy, ValueEnum)]
enum StrategyArg {
    Fail,
    Ours,
    Theirs,
}

impl From<StrategyArg> for MergeStrategy {
    fn from(value: StrategyArg) -> Self {
        match value {
            StrategyArg::Fail => MergeStrategy::Fail,
            StrategyArg::Ours => MergeStrategy::Ours,
            StrategyArg::Theirs => MergeStrategy::Theirs,
        }
    }
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
            labels,
            layout,
        } => {
            let result = diff_files(&old, &new)?;
            match format {
                OutputFormat::Text => println!("{}", format_text_diff(&result)),
                OutputFormat::Json => println!("{}", serde_json::to_string_pretty(&result)?),
                OutputFormat::Excel => {
                    let output = output.unwrap_or_else(|| default_diff_path(&new));
                    let labels = labels
                        .as_deref()
                        .map(DiffReportLabels::from_json_path)
                        .transpose()?
                        .unwrap_or_default();
                    write_diff_excel_with_layout(&result, &output, &labels, layout.into())?;
                    println!("{}", output.display());
                }
            }
            Ok(if result.issues.iter().any(|issue| issue.is_error()) {
                1
            } else {
                0
            })
        }
        Command::Merge {
            base,
            ours,
            theirs,
            strategy,
            output,
            report,
            allow_conflicts,
        } => {
            let result = merge_files(&base, &ours, &theirs, strategy.into())?;
            match report {
                ReportFormat::Text => println!("{}", format_merge_text(&result)),
                ReportFormat::Json => {
                    println!(
                        "{}",
                        serde_json::to_string_pretty(&json!({
                            "ok": result.ok(),
                            "strategy": result.strategy,
                            "summary": result.summary(),
                            "changes": &result.changes,
                            "issues": &result.issues,
                        }))?
                    );
                }
            }
            if result.has_unresolved_conflicts() && !allow_conflicts {
                eprintln!(
                    "merge not written: {} unresolved conflicts",
                    result.conflict_count()
                );
                return Ok(2);
            }
            let output = output.unwrap_or_else(|| default_merge_path(&ours));
            write_merge_excel(&result, &output, true)?;
            println!("{}", output.display());
            Ok(if result.ok() { 0 } else { 2 })
        }
        Command::InstallSvn { config, dry_run } => {
            let executable = std::env::current_exe()?;
            let result = install_svn_integration(executable, config.as_deref(), dry_run)?;
            if result.dry_run {
                println!("SVN 配置路径: {}", result.config_path);
                println!("diff-cmd: {}", result.diff_cmd);
                println!("merge-tool-cmd: {}", result.merge_tool_cmd);
                println!("仅预览，未写入配置。");
            } else if result.changed {
                println!("已注册 TableManager SVN 集成: {}", result.config_path);
                println!("diff-cmd: {}", result.diff_cmd);
                println!("merge-tool-cmd: {}", result.merge_tool_cmd);
            } else {
                println!("SVN 集成已指向当前 TableManager: {}", result.config_path);
            }
            Ok(0)
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

fn default_merge_path(path: &Path) -> PathBuf {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("workbook");
    path.with_file_name(format!("{stem}_merged.xlsx"))
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
