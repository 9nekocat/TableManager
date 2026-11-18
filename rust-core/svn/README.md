# SVN Excel 比对与三路合并

包装脚本把 Subversion 和 TortoiseSVN 接到 Rust `table-core`。Excel 文件使用表级比对与合并，其他文件仍交给已安装的 xlCompare。

以下命令以发布包解压到 `C:\Tools\table-core` 为例。

## Subversion 命令行

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

- 差异会生成 Summary、Changes、Issues 工作簿并打开审阅。
- 合并会生成 `_MergeReview`、`_MergeConflicts`、`_MergeChanges` 和 `_MergeIssues`。
- 存在未解决冲突时仍写出审查工作簿，但返回退出码 `2`，Subversion 不会将其视为成功解决。
- 支持 `.xlsx.r1234`、`.xlsx.mine` 等 SVN 临时文件名。

环境变量：

```text
TABLE_CORE_EXE       指定 table-core.exe 路径
TM_SVN_NO_OPEN=1     只生成文件，不打开 Excel
TM_SVN_OUTPUT        指定 diff 输出文件，主要用于自动化测试
```
