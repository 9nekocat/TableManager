# Merge Post Cleanup

合并成功后的冲突产物文件清理。

## ADDED Requirements

### Requirement: Cleanup SVN conflict artifacts after successful merge

合并成功后，系统 SHALL 删除 SVN 冲突产生的临时文件（`.mine`、`.r*` 等）。

#### Scenario: Web API deletes r-files after merge success
- **WHEN** `api_merge_resolve` 合并成功后执行 `svn resolved`
- **THEN** 系统删除 `RAW_DIR` 中匹配 `{stem}.*.r*` 和 `{stem}.xlsx.mine` 的所有文件

#### Scenario: CLI deletes r-files after merge success
- **WHEN** `cmd_merge_resolve` 合并成功后执行 `svn resolved`
- **THEN** 系统删除 `RAW_DIR` 中匹配 `{stem}.*.r*` 和 `{stem}.xlsx.mine` 的所有文件

#### Scenario: Cleanup failure does not block merge result
- **WHEN** 清理过程中个别文件删除失败（如被其他进程占用）
- **THEN** 系统记录警告但不影响合并成功的返回值

### Requirement: Shared cleanup function

清理逻辑 SHALL 提取为独立函数供 CLI 和 Web API 复用。

#### Scenario: Cleanup function is callable from both entry points
- **WHEN** CLI 或 Web API 需要清理冲突文件
- **THEN** 调用 `cleanup_conflict_artifacts(stem, raw_dir)` 执行清理，返回已删除和失败的列表
