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
- 合并会生成 `_MergeReview`、`_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues`。
- 存在未解决冲突时仍写出审查工作簿，但返回退出码 `2`，Subversion 不会将其视为成功解决。
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
  "remember_output_dir": true
}
```

首选目录写入失败时会自动尝试回退目录。`remember_output_dir` 为 `true` 时，成功的
回退路径会写回配置，后续比对直接使用该路径。`${SOURCE_DIR}` 表示当前被比对工作表
所在目录。也可以把 `diff_output_dir` 改为任意绝对路径。
