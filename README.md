# TableManager

轻量、无界面的 Excel 配置表比对与三路合并工具。核心由 Rust 实现，面向人类命令行用户、AI 自动化和 SVN/TortoiseSVN。

## 功能

- 两文件比对，支持类 Git 文本、JSON 和可直接查看的 Excel 报告。
- 三路合并，支持单元格、行、列和工作表的增删改。
- 冲突策略：`fail`、`ours`、`theirs`；默认 `fail`。
- 合并结果附带 Review、Conflicts、Changes 和 Issues 审查页。
- 保留简单公式，并在行列移动后重绑定行内引用。
- 识别重复表头和 RowName、复杂公式风险、异常空行空列。
- 非公式但以 `=` 开头的文本会按 Excel 文本格式写回。
- 提供 SVN 命令行和 TortoiseSVN 的 Excel 比对/合并包装。

## 下载与启动

从 GitHub Releases 下载 `table-core-v0.1.0-windows-x64.zip`，解压后运行：

```powershell
.\table-core.exe --help
```

无需 Python、Node.js 或 Rust 运行时。

## 命令行

```powershell
# 类 Git 文本差异
.\table-core.exe diff old.xlsx new.xlsx --format text

# JSON，适合 AI 或自动化
.\table-core.exe diff old.xlsx new.xlsx --format json

# Excel 差异报告
.\table-core.exe diff old.xlsx new.xlsx --format excel --output diff.xlsx

# 三路合并
.\table-core.exe merge base.xlsx ours.xlsx theirs.xlsx --strategy fail --output merged.xlsx
```

冲突存在时，可加 `--allow-conflicts` 写出带 `_MergeConflicts` 页的审查文件。退出码：

- `0`：执行成功，没有未解决冲突。
- `1`：比较已完成，但报告中包含错误级问题。
- `2`：存在未解决冲突，或文件生成失败。

## Excel 约定

- 只有名称以 `#` 开头的 sheet 作为数据页。
- 第 1、2、3 行分别是中文字段名、英文字段名和类型。
- 数据从第 4 行开始。
- A 列是 RowName；重复 RowName 会保留并以稳定序号区分，同时给出警告。
- 比对时公式使用 Excel 中已保存的缓存值。
- 合并支持行内引用和固定目标跨页引用；跨行相对引用、外部工作簿、动态数组等复杂公式会进入风险报告。

## SVN 配置

发布包中的 `svn` 目录包含包装脚本。Subversion 命令行配置：

```ini
[helpers]
diff-cmd = C:/path/to/table-core/svn/svn-diff.cmd
diff-extensions =
merge-tool-cmd = C:/path/to/table-core/svn/svn-merge.cmd

[miscellany]
diff-ignore-content-type = yes
```

TortoiseSVN：

```text
Diff:  "C:\path\to\table-core\svn\svn-diff.cmd" %base %mine
Merge: "C:\path\to\table-core\svn\svn-merge.cmd" %merged %theirs %mine %base
```

SVN 合并存在冲突时包装脚本返回 `2`，不会把冲突误报为已解决。

## 从源码构建

需要 Rust stable：

```powershell
cargo build --release --locked --manifest-path rust-core/Cargo.toml
```

构建 Windows 发布包：

```powershell
powershell -ExecutionPolicy Bypass -File rust-core/scripts/package-release.ps1 -Version 0.1.0
```

产物位于 `rust-core/dist`。

## 验证

```powershell
cargo test --locked --manifest-path rust-core/Cargo.toml
cargo clippy --locked --all-targets --manifest-path rust-core/Cargo.toml -- -D warnings
```
