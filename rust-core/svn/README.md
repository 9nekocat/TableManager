# SVN Excel 比对与三路合并

包装脚本用 Rust `table-core` 对 `.xlsx`、`.xlsm`、`.xlsb`、`.xls` 进行表级比对与三路合并。

以下命令以发布包解压到 `C:\Tools\table-core` 为例。

## TortoiseSVN 注册

根目录的 `configure-svn.bat` 和 `table-core.exe install-svn` 都可以注册四种 Excel 扩展名。
注册表备份位于 `%LOCALAPPDATA%\CatsInTheShell\TableManager\backup`。

```powershell
.\configure-svn.bat
.\configure-svn-管理员模式运行.bat

C:\Tools\table-core\table-core.exe install-svn --dry-run
C:\Tools\table-core\table-core.exe install-svn
```

手动注册可以在 **Settings > Diff Viewer > Advanced** 和 **Settings > Merge Tool > Advanced**
为这四种扩展名指定以下命令。

```text
Diff:  "C:\Tools\table-core\svn\svn-diff.cmd" %base %mine
Merge: "C:\Tools\table-core\svn\svn-merge.cmd" %merged %theirs %mine %base
```

## 行为

- xlsx 族差异会生成排列模式报告：每个原表页签独立，最后是比对报告。
- 无真实冲突时自动生成干净的业务工作簿，并执行 `svn resolve --accept working`。
- 有真实冲突时生成三路合并预览：每张原表页签按 `[基线]`、`[我的]`、`[仓库]`、
  `[最终]` 排列。`[最终]` 行位于每组最下方，并预填配置指定的默认侧。预览默认保存到
  TableManager 目录下的 `.temp`。
- 预览工作簿以只读方式打开，命令窗口立即等待用户输入 `Y` 确认；用户无需先关闭
  Excel，由用户自行决定查看完成时间。确认后读取 `[最终]` 行写回，并执行
  `svn resolve --accept working`。
- 非无头模式会在合并完成或取消后显示一个 15 秒的 TableManager 提示框，
  避免从 TortoiseSVN 启动时命令窗口一闪而过。
- 单元格冲突默认使用仓库值 `theirs`；可在 `svn-wrapper.json` 的
  `merge_conflict_default` 中改为 `ours` 或 `empty`。
- `_MergeMeta` 是隐藏回写元数据页；正式业务工作簿不会包含任何 `_Merge*` 或预览页签。
- 支持 `.xlsx.r1234`、`.xlsx.mine` 等 SVN 临时文件名。

## 无头模式

`svn-diff-headless.cmd` 和 `svn-merge-headless.cmd` 设置 `TM_SVN_HEADLESS=1`。
在命令行中，可直接调用这两个包装器来生成表格报告。包装器遇到不受支持的文件类型会返回退出码 `2`。

环境变量：

```text
TABLE_CORE_EXE       指定 table-core.exe 路径
TABLE_CORE_SVN_DIR   指定包含 configure-tortoisesvn.ps1 及包装脚本的目录
TM_SVN_NO_OPEN=1     只生成文件，不打开 Excel
TM_SVN_HEADLESS=1    无头模式；不打开 Excel，也不回退到外部工具
TM_SVN_OUTPUT        指定 diff 输出文件，主要用于自动化测试
TM_SVN_OUTPUT_DIR    指定 diff 首选输出目录
TM_SVN_CONFIG        指定包装器配置文件；默认使用用户级 config\svn-wrapper.json
TM_SVN_MERGE_REVIEW  指定冲突预览文件路径，主要用于自动化测试
TM_SVN_NO_TIP=1      不显示合并结束提示框
```

发布包中的 `config\svn-wrapper.json` 是默认模板。实际记忆路径为：

```text
%LOCALAPPDATA%\CatsInTheShell\TableManager\config\svn-wrapper.json
```

配置文件控制输出策略：

```json
{
  "diff_output_dir": "%TEMP%\\TableManager\\SVN",
  "fallback_diff_output_dir": "${SOURCE_DIR}\\.temp",
  "remember_output_dir": true,
  "merge_conflict_default": "theirs"
}
```

首选目录写入失败时会自动尝试配置中的回退目录，再尝试 TableManager 目录下的 `.temp`。
`remember_output_dir` 为 `true` 时，配置中的回退目录写入成功后会被记为首选目录。
`${SOURCE_DIR}` 表示当前被比对工作表所在目录，在 TortoiseSVN 日志中比对时可能指向
历史版本的临时目录。也可以把 `diff_output_dir` 改为任意绝对路径。

`merge_conflict_default` 支持：

- `theirs`：默认采用仓库值（推荐）
- `ours`：默认采用本地值
- `empty`：冲突单元格默认留空
