# SVN Excel 比对与三路合并

包装脚本把 Subversion 和 TortoiseSVN 接到 Rust `table-core`。`.xlsx`、`.xlsm`、
`.xlsb`、`.xls` 全部由 TableManager 自身完成表级比对与三路合并；普通文本文件仍可
交给已安装的 xlCompare。

以下命令以发布包解压到 `C:\Tools\table-core` 为例。

## Subversion 命令行

首选自动注册。根目录的 `configure-svn.bat` 会同时配置：

- 当前用户的 Subversion `diff-cmd` / `merge-tool-cmd`；
- TortoiseSVN 的 `.xlsx`、`.xlsm`、`.xlsb`、`.xls` DiffTools 和 MergeTools。

Subversion 配置会保留 `config.table-core.bak`；TortoiseSVN 注册表会备份到
`%LOCALAPPDATA%\CatsInTheShell\TableManager\backup`。

```powershell
.\configure-svn.bat
.\configure-svn-管理员模式运行.bat

C:\Tools\table-core\table-core.exe install-svn --dry-run
C:\Tools\table-core\table-core.exe install-svn
```

`table-core.exe install-svn` 只修改 Subversion 命令行配置；需要同时配置 TortoiseSVN 时，
使用根目录的一键 BAT。

编辑 `%APPDATA%\Subversion\config`：

```ini
[helpers]
diff-cmd = C:/Tools/table-core/svn/svn-diff.cmd
diff-extensions =
merge-tool-cmd = C:/Tools/table-core/svn/svn-merge.cmd

[miscellany]
diff-ignore-content-type = yes
```

## TortoiseSVN

在 **Settings > Diff Viewer > Advanced** 和 **Settings > Merge Tool > Advanced** 中配置 `.xlsx`、`.xlsm`、`.xlsb`、`.xls`：

```text
Diff:  "C:\Tools\table-core\svn\svn-diff.cmd" %base %mine
Merge: "C:\Tools\table-core\svn\svn-merge.cmd" %merged %theirs %mine %base
```

## 行为

- xlsx 族差异会生成排列模式报告：每个原表页签独立，最后是比对报告。
- 无真实冲突时自动生成干净的业务工作簿，并执行 `svn resolve --accept working`。
- 有真实冲突时生成三路合并预览：每张原表页签按 `[基线]`、`[我的]`、`[仓库]`、
  `[最终]` 排列。`[最终]` 行位于每组最下方，并预填配置指定的默认侧。
- 预览关闭后，命令窗口会要求用户输入 `Y` 确认，再读取 `[最终]` 行写回，
  最后执行 `svn resolve --accept working`。
- 非无头模式会在合并完成或取消后显示一个 15 秒的 TableManager 提示框，
  避免从 TortoiseSVN 启动时命令窗口一闪而过。
- 单元格冲突默认使用仓库值 `theirs`；可在 `svn-wrapper.json` 的
  `merge_conflict_default` 中改为 `ours` 或 `empty`。
- `_MergeMeta` 是隐藏回写元数据页；正式业务工作簿不会包含任何 `_Merge*` 或预览页签。
- 支持 `.xlsx.r1234`、`.xlsx.mine` 等 SVN 临时文件名。

## 无头模式

`svn-diff-headless.cmd` 和 `svn-merge-headless.cmd` 设置 `TM_SVN_HEADLESS=1`，
不打开 Excel，也不回退到 xlCompare。未识别的格式会明确返回退出码 `2`。

环境变量：

```text
TABLE_CORE_EXE       指定 table-core.exe 路径
TABLE_CORE_SVN_DIR   指定包含 svn-diff.cmd / svn-merge.cmd 的目录
TM_SVN_NO_OPEN=1     只生成文件，不打开 Excel
TM_SVN_HEADLESS=1    无头模式；不打开 Excel，也不回退到外部工具
TM_SVN_OUTPUT        指定 diff 输出文件，主要用于自动化测试
TM_SVN_OUTPUT_DIR    指定 diff 输出目录；默认使用工作文件旁边的 .temp
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

首选目录写入失败时会自动尝试回退目录。`remember_output_dir` 为 `true` 时，成功的
回退路径会写回配置，后续比对直接使用该路径。`${SOURCE_DIR}` 表示当前被比对工作表
所在目录。也可以把 `diff_output_dir` 改为任意绝对路径。

`merge_conflict_default` 支持：

- `theirs`：默认采用仓库值（推荐）
- `ours`：默认采用本地值
- `empty`：冲突单元格默认留空
