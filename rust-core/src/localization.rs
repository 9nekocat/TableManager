use std::path::Path;

use serde::{Deserialize, Serialize};

use crate::error::Result;

/// User-facing text used by the Excel diff report.
///
/// The defaults are Simplified Chinese. Settings can provide a partial JSON
/// object to override any label while keeping the remaining defaults.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct DiffReportLabels {
    pub report_title: String,
    pub summary: String,
    pub old_file: String,
    pub new_file: String,
    pub comparison: String,
    pub comparison_cached_values: String,
    pub has_changes: String,
    pub rows_added: String,
    pub rows_deleted: String,
    pub rows_modified: String,
    pub columns_added: String,
    pub columns_deleted: String,
    pub columns_modified: String,
    pub columns_moved: String,
    pub cells_modified: String,
    pub yes: String,
    pub no: String,
    pub no_differences: String,
    pub worksheet: String,
    pub chinese_header: String,
    pub english_header: String,
    pub field_type: String,
    pub old_row: String,
    pub new_row: String,
    pub added_row: String,
    pub deleted_row: String,
    pub added_column: String,
    pub deleted_column: String,
    pub modified_column: String,
    pub moved_column: String,
    pub issues: String,
    pub severity: String,
    pub code: String,
    pub sheet: String,
    pub row: String,
    pub column: String,
    pub message: String,
    pub reference: String,
    pub blank_rows_removed_group: String,
    pub error: String,
    pub warning: String,
    pub info: String,
}

impl Default for DiffReportLabels {
    fn default() -> Self {
        Self {
            report_title: "比对报告".to_string(),
            summary: "摘要".to_string(),
            old_file: "旧文件".to_string(),
            new_file: "新文件".to_string(),
            comparison: "比对方式".to_string(),
            comparison_cached_values: "基于已保存的公式缓存值".to_string(),
            has_changes: "存在差异".to_string(),
            rows_added: "新增行".to_string(),
            rows_deleted: "删除行".to_string(),
            rows_modified: "修改行".to_string(),
            columns_added: "新增列".to_string(),
            columns_deleted: "删除列".to_string(),
            columns_modified: "修改列".to_string(),
            columns_moved: "移动列".to_string(),
            cells_modified: "修改单元格".to_string(),
            yes: "是".to_string(),
            no: "否".to_string(),
            no_differences: "未发现差异".to_string(),
            worksheet: "工作表".to_string(),
            chinese_header: "中文字段名".to_string(),
            english_header: "英文字段名".to_string(),
            field_type: "字段类型".to_string(),
            old_row: "旧版".to_string(),
            new_row: "当前".to_string(),
            added_row: "新增".to_string(),
            deleted_row: "删除".to_string(),
            added_column: "新增列".to_string(),
            deleted_column: "删除列".to_string(),
            modified_column: "修改列".to_string(),
            moved_column: "移动列".to_string(),
            issues: "检查问题".to_string(),
            severity: "级别".to_string(),
            code: "代码".to_string(),
            sheet: "工作表".to_string(),
            row: "行".to_string(),
            column: "列".to_string(),
            message: "说明".to_string(),
            reference: "引用".to_string(),
            blank_rows_removed_group: "检测到 {count} 个空数据行，加载时已从模型中移除。"
                .to_string(),
            error: "错误".to_string(),
            warning: "警告".to_string(),
            info: "提示".to_string(),
        }
    }
}

impl DiffReportLabels {
    pub fn from_json_path(path: impl AsRef<Path>) -> Result<Self> {
        let contents = std::fs::read_to_string(path)?;
        Ok(serde_json::from_str(&contents)?)
    }
}
