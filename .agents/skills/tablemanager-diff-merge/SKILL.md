---
name: tablemanager-diff-merge
description: "用于 catsintheshell/TableManager 的 Excel 配置表比对与三路合并：AI 自查必须使用无头纯文本，用户审阅必须生成排列式 Excel。"
---

# TableManager 比对与合并

把“机器检查”和“人看结果”明确分开。两者不能混用。

## 先判断受众

### AI 自查、验证、定位冲突

必须使用无头纯文本模式：

- 不启动 Excel、LibreOffice、xlCompare 或任何 GUI。
- 不生成 Excel 报告作为 AI 的第一手检查材料。
- 使用 `table-core.exe diff ... --format text`。
- 检查三路合并时使用 `table-core.exe merge ... --report text`；退出码 `2` 表示存在真实冲突。
- 只把文本摘要、退出码和必要的单元格引用带回推理过程。
- `svn-diff-headless.cmd` 仍会生成 Excel，不能替代纯文本 AI 自查。

### 给用户展示、让用户审阅

必须使用 Excel，并优先使用排列式报告：

- 双文件比对使用 `--format excel --layout arranged`。
- SVN 比对优先走 `svn/svn-diff.cmd` 或已配置的 `svn diff`，让输出目录和自动打开策略保持项目一致。
- 三路合并优先走 `svn/svn-merge.cmd`，由它生成只读三路预览并等待用户在命令行确认。
- 不要把原始文本 diff 当作最终用户界面；文本只用于 AI 自查或命令行诊断。

## 固定工作流

1. 先做 AI 文本自查，确认改动范围、公式风险、重复行/列和冲突数量。
2. 自查通过后，再为用户生成 Excel 报告或打开三路合并预览。
3. 如果文本自查发现真实冲突，不得自行采用 `ours`/`theirs` 掩盖冲突，也不得提前执行
   `svn resolve`。等待用户确认最终值。
4. 只有最终表已写回且用户确认后，才允许执行 `svn resolve --accept working`。

## 退出码

- `0`：比对完成且没有错误级问题；或合并没有真实冲突。
- `1`：比对报告包含错误级检查项，需要先处理。
- `2`：存在未解决冲突，或命令执行失败，不能提交为已解决。

## 路径

从项目根目录运行：

```text
Tools/TableManager/table-core.exe
Tools/TableManager/svn/svn-diff.cmd
Tools/TableManager/svn/svn-merge.cmd
```

当前目录已经是 TableManager 时使用：

```text
.\table-core.exe
.\svn\svn-diff.cmd
.\svn\svn-merge.cmd
```

具体命令矩阵、配置项和失败处理见
[references/commands.md](references/commands.md)。需要执行比对、合并或排查 SVN 包装器时读取该参考。
