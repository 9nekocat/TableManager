## ADDED Requirements

### Requirement: Diff export writes to .temp directory

`export_diff_to_excel` SHALL 将 diff Excel 文件输出到 `RAW_DIR/.temp/` 目录下。

#### Scenario: Export diff for a file

- **WHEN** 对 `Effect.xlsx` 执行 diff 导出
- **THEN** 输出文件路径 SHALL 为 `RAW_DIR/.temp/Effect_diff_r{rev}.xlsx`

### Requirement: Merge preview writes to .temp directory

`api_merge_preview` SHALL 将 merge preview Excel 输出到 `RAW_DIR/.temp/` 目录下。

#### Scenario: Generate merge preview

- **WHEN** 对冲突文件生成合并预览
- **THEN** 预览文件路径 SHALL 为 `RAW_DIR/.temp/{stem}_merge.xlsx`

### Requirement: Merge resolve reads from .temp directory

`api_merge_resolve` SHALL 从 `RAW_DIR/.temp/` 读取 merge preview 文件。

#### Scenario: Resolve merge from preview

- **WHEN** 执行合并解决
- **THEN** 系统 SHALL 从 `RAW_DIR/.temp/` 查找 preview 文件，并正确推断原文件路径（基于 RAW_DIR 而非字符串替换）

### Requirement: Merge abort reads and cleans from .temp directory

`api_merge_abort` SHALL 从 `RAW_DIR/.temp/` 读取并删除 merge preview 文件。

#### Scenario: Abort merge

- **WHEN** 用户放弃合并
- **THEN** 系统 SHALL 删除 `RAW_DIR/.temp/{stem}_merge.xlsx`
