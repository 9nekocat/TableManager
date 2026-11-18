# Diff No-Recalc Optimized

在文件已有缓存公式值的情况下跳过 Excel recalc，直接从 xlsx 读取。

## ADDED Requirements

### Requirement: Skip recalc for SVN-exported old versions

从 SVN 导出的历史版本 SHALL 不经过 Excel recalc 直接读取。SVN 仓库中的 xlsx 由用户通过 Excel 提交，已有完整缓存值。

#### Scenario: Old version read without recalc
- **WHEN** `_get_old_version("Effect.xlsx", "HEAD")` 执行
- **THEN** SVN cat 导出后直接从 temp 文件读取，不调用 Excel recalc

### Requirement: Skip recalc for working copy in generate_diff

`generate_diff()` 和 `export_diff_to_excel()` SHALL 直接从工作副本读取当前版本，不走 recalc_to_temp 复制+重算流程。

#### Scenario: Current version read directly
- **WHEN** 工作副本 `Effect.xlsx` 由 Excel 保存且未被 TableManager 修改
- **THEN** `_read_all_rows_with_formulas` 直接读取文件，公式缓存值可用

### Requirement: Batch recalc for diff_two_files

`diff_two_files()` SHALL 在需要保证公式缓存值时使用 copy-to-temp + `recalc_batch` 策略。

#### Scenario: Two-file diff with guaranteed cached values
- **WHEN** `diff_two_files(old_path, new_path)` 执行
- **THEN** 两个文件被复制到临时位置，批量重算，读取后删除临时副本
- **THEN** diff 输出包含实际计算值而非空白
