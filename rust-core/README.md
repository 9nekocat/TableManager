# TableManager

Rust Excel 配置表比对与三路合并命令行工具。

```powershell
table-core.exe diff OLD NEW --format text
table-core.exe diff OLD NEW --format json
table-core.exe diff OLD NEW --format excel --output diff.xlsx
table-core.exe merge BASE OURS THEIRS --strategy fail --output merged.xlsx
```

主要行为：

- 公式比对使用 Excel 缓存值。
- 合并保留简单公式，并在行列变更后重绑定行内引用。
- 默认冲突策略为 `fail`；未解决冲突返回退出码 `2`。
- `--allow-conflicts` 写出 `_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues` 审查页。
- 复杂公式、重复表头、重复 RowName 和空行列会进入检查报告。

SVN 配置见 [svn/README.md](svn/README.md)。
