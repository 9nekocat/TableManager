# TableManager

Rust Excel 配置表比对与三路合并命令行工具。

```powershell
table-core.exe diff OLD NEW --format text
table-core.exe diff OLD NEW --format json
table-core.exe diff OLD NEW --format excel --output diff.xlsx
table-core.exe diff OLD NEW --format excel --layout compact --output diff.xlsx
table-core.exe diff OLD NEW --format excel --output diff.xlsx --labels config\diff-report.zh-CN.json
table-core.exe merge BASE OURS THEIRS --strategy fail --conflict-default theirs --output merged.xlsx
table-core.exe merge BASE OURS THEIRS --strategy fail --conflict-default theirs --review-output review.xlsx
table-core.exe resolve-review review.xlsx BASE OURS THEIRS --conflict-default theirs --output merged.xlsx
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
- 默认冲突策略为 `fail`；真实冲突默认预填仓库值 `theirs`，可用
  `--conflict-default ours|theirs|empty` 修改。
- 合并预览按原表页签生成 `[基线]`、`[我的]`、`[仓库]`、`[最终]` 行；用户最终采用值放在
  每条冲突记录最下方的 `[最终]` 行。
- `resolve-review` 读取 `[最终]` 行并写出干净的业务工作簿，不会把预览页签带入正式表。
- SVN 合并没有真实冲突时会自动写出结果并执行 `svn resolve --accept working`；有真实冲突时
  以只读方式打开预览，并在命令行等待用户确认后再写回和标记 resolved，用户无需先关闭 Excel。
- 非无头模式在合并完成或取消后会显示 TableManager 提示框，避免命令窗口自动关闭后
  用户看不到结果；`TM_SVN_NO_TIP=1` 可关闭提示框。
- `--allow-conflicts` 写出 `_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues` 审查页。
- 复杂公式、重复表头、重复 RowName 和空行列会进入检查报告。

SVN 配置见 [svn/README.md](svn/README.md)。
