# SVN Excel 比对与三路合并

包装脚本把 Subversion 和 TortoiseSVN 接到 Rust `table-core`。`.xlsx`、`.xlsm`、
`.xlsb`、`.xls` 全部由 TableManager 自身完成表级比对与三路合并；普通文本文件仍可
交给已安装的 xlCompare。

以下命令以发布包解压到 `C:\Tools\table-core` 为例。

## Subversion 命令行

首选自动注册（写入当前用户的 Subversion 配置，并保留 `config.table-core.bak`）：

```powershell
C:\Tools\table-core\table-core.exe install-svn --dry-run
C:\Tools\table-core\table-core.exe install-svn
```

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

- xlsx 族差异会生成单页比对报告并打开审阅。
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
```
