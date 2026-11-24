# TableManager

TableManager 帮助游戏策划在版本管理中协作编写 Excel 配置表。提交前可以比对两个版本，按记录和字段查看改动，例如哪件道具的售价从 50 改到了 60。多人编辑同一张表时，工具可以用共同基线、本地版本和仓库版本做三路合并，自动合并互不冲突的改动，把双方对同一字段的不同修改列出来供策划确认。

## 配置表格式

参与比对与合并的工作表名必须以 `#` 开头，例如 `#道具`。每张工作表的前 3 行是表头，从第 4 行起每行一条记录。

| Excel 行 | A 列 | B 列 | C 列 |
| --- | --- | --- | --- |
| 1 · 中文字段名 | 道具编号 | 名称 | 售价 |
| 2 · 英文字段名 | ItemID | Name | Price |
| 3 · 字段类型 | string | string | int |
| 4 · 第一条记录 | potion_hp | 生命药水 | 50 |
| 5 · 第二条记录 | potion_mp | 法力药水 | 80 |

工具优先按第 2 行的英文字段名匹配列。请为每列填写稳定、在本表内唯一的英文字段名，匹配时忽略大小写。

### 设置 Key 列

把记录的唯一标识放在**最左侧有内容的列**（通常是 A 列）。工具按位置取这一列作为行 Key（RowName），无需额外配置，字段名可以使用示例中的 `ItemID`。Key 前面的整列如果完全空白，会被跳过。

同一工作表内，每条记录的 Key 应非空、唯一，并在各版本中保持稳定。例如，把 `potion_hp` 移到另一行后，工具仍能识别同一件道具。若将它改名为一个新 Key，比对会显示旧记录删除、新记录增加。缺少 Key 时工具会生成临时行键并给出警告，重复 Key 会报错。

### 公式限制

**不支持复杂公式。** 比对读取 Excel 上次保存的计算结果。提交前请在 Excel 中重新计算并保存，缺少缓存结果的公式会按空值参与比对，并产生警告。

合并可以保留简单的行内公式，例如在 D4 写 `=C4*2`，记录移动时会调整相对引用。跨工作表引用须固定行和列，例如 `='#参数'!$B$4`。同表跨数据行引用、跨工作簿引用会报错，`INDIRECT`、`FILTER`、`XLOOKUP` 等复杂函数可能触发警告。

## 使用方法

解压发布包后，在该目录打开 PowerShell。`OLD` 和 `NEW` 是要比对的两个文件路径，`BASE`、`OURS`、`THEIRS` 分别是共同基线、本地文件和仓库文件的路径。

```powershell
.\table-core.exe diff OLD NEW --format text
.\table-core.exe diff OLD NEW --format json
.\table-core.exe diff OLD NEW --format excel --output diff.xlsx
.\table-core.exe diff OLD NEW --format excel --layout compact --output diff.xlsx
.\table-core.exe diff OLD NEW --format excel --output diff.xlsx --labels config\diff-report.zh-CN.json
.\table-core.exe merge BASE OURS THEIRS --strategy fail --conflict-default theirs --output merged.xlsx
.\table-core.exe merge BASE OURS THEIRS --strategy fail --conflict-default theirs --review-output review.xlsx
.\table-core.exe resolve-review review.xlsx BASE OURS THEIRS --conflict-default theirs --output merged.xlsx
```

## 比对报告与合并审阅

- Excel 报告默认使用 `--layout arranged`，每张原表有独立页签，沿用原表列头，最后一个页签为“比对报告”。
- `--layout compact` 将摘要和变更放在同一页。
- 排列模式会合并同类空行问题，错误级问题和其他重要问题逐条说明。
- 报告辅助文案由 `DiffReportLabels` 本地化配置控制，JSON 可只覆盖需要修改的字段。
- 默认冲突策略为 `fail`，真实冲突默认预填仓库值 `theirs`，可用
  `--conflict-default ours|theirs|empty` 修改。
- 合并预览按原表页签生成 `[基线]`、`[我的]`、`[仓库]`、`[最终]` 行。用户最终采用值放在
  每条冲突记录最下方的 `[最终]` 行。
- `resolve-review` 读取 `[最终]` 行并写出干净的业务工作簿，不会把预览页签带入正式表。
- `--allow-conflicts` 写出 `_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues` 审查页。

## 接入 TortoiseSVN

双击发布包中的 `configure-svn.bat`，即可为当前用户注册 `.xlsx`、`.xlsm`、`.xlsb`、`.xls` 的比对与合并工具。管理员入口为 `configure-svn-管理员模式运行.bat`。命令行也可用 `.\table-core.exe install-svn --dry-run` 预览注册内容，再用 `.\table-core.exe install-svn` 执行注册。

- SVN 合并没有真实冲突时会自动写出结果并执行 `svn resolve --accept working`。有真实冲突时
  以只读方式打开预览，并在命令行等待用户确认后再写回和标记 resolved，用户无需先关闭 Excel。
- 交互模式在合并完成或取消后会显示 TableManager 提示框，`TM_SVN_NO_TIP=1` 可关闭提示框。
- 自动化流程可以调用 `svn-diff-headless.cmd` 和 `svn-merge-headless.cmd` 在后台运行，无需打开 Excel。

详细配置见发布包中的 `svn/README.md`（源码仓库中位于 `rust-core/svn/README.md`）。
