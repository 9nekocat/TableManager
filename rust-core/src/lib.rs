pub mod diff;
pub mod error;
pub mod formula;
pub mod localization;
pub mod merge;
pub mod merge_review;
pub mod model;
pub mod output;
pub mod snapshot;
pub mod svn_integration;

pub use diff::{
    CellChange, ColumnAlignment, ColumnChange, DiffResult, RowChange, SheetDiff, diff_files,
    diff_workbooks, format_text_diff,
};
pub use error::{CoreError, Result};
pub use formula::{FormulaFinding, analyze_formula};
pub use localization::DiffReportLabels;
pub use merge::{
    ConflictDefault, MergeChange, MergePreviewRow, MergeResult, MergeStrategy, MergeSummary,
    format_merge_text, merge_files, merge_files_with_conflict_default, merge_workbooks,
    merge_workbooks_with_conflict_default,
};
pub use merge_review::apply_merge_review;
pub use model::{Cell, Header, Issue, Row, Sheet, Workbook};
pub use output::{DiffExcelLayout, write_merge_review, write_merged_workbook};
pub use snapshot::{TextRow, TextSheet, TextWorkbook, workbook_to_text};
pub use svn_integration::install_svn_integration;
