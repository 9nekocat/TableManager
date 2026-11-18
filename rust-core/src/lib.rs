pub mod diff;
pub mod error;
pub mod formula;
pub mod localization;
pub mod merge;
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
    MergeChange, MergeResult, MergeStrategy, MergeSummary, format_merge_text, merge_files,
    merge_workbooks,
};
pub use model::{Cell, Header, Issue, Row, Sheet, Workbook};
pub use output::DiffExcelLayout;
pub use snapshot::{TextRow, TextSheet, TextWorkbook, workbook_to_text};
pub use svn_integration::{SvnInstallResult, install_svn_integration};
