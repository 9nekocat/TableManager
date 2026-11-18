# TableManager

Rust Excel 配置表比对与三路合并命令行工具。

```powershell
table-core.exe diff OLD NEW --format text
table-core.exe diff OLD NEW --format json
table-core.exe diff OLD NEW --format excel --output diff.xlsx
table-core.exe diff OLD NEW --format excel --layout compact --output diff.xlsx
table-core.exe diff OLD NEW --format excel --output diff.xlsx --labels config\diff-report.zh-CN.json
table-core.exe merge BASE OURS THEIRS --strategy fail --output merged.xlsx
configure-svn.bat
table-core.exe install-svn --dry-run
table-core.exe install-svn
```

主要行为：

- 公式比对使用 Excel 缓存值。
- Excel 比对报告将摘要和变更放在同一页；变更区沿用原表列头和单元格布局。
- Excel 默认使用 `--layout arranged`：每个原表页签独立，最后一个页签为比对报告。
- `--layout compact` 可选择原有的紧凑单页样式。
- 排列模式会合并同类空行问题；错误级问题和其他重要问题逐条说明。
- 报告辅助文案由 `DiffReportLabels` 本地化配置控制，JSON 可只覆盖需要修改的字段。
- `install-svn` 会把当前用户的 Subversion `diff-cmd` / `merge-tool-cmd` 指向发布包自身。
- `configure-svn.bat` 是一键注册入口，同时配置 Subversion 命令行和 TortoiseSVN；双击或在当前目录运行即可。
- `configure-svn-管理员模式运行.bat` 是同一流程的管理员模式入口。
- 发布包包含 `svn-diff-headless.cmd` 和 `svn-merge-headless.cmd`，不会打开 Excel。
- 合并保留简单公式，并在行列变更后重绑定行内引用。
- 默认冲突策略为 `fail`；未解决冲突返回退出码 `2`。
- `--allow-conflicts` 写出 `_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues` 审查页。
- 复杂公式、重复表头、重复 RowName 和空行列会进入检查报告。

SVN 配置见 [svn/README.md](svn/README.md)。
