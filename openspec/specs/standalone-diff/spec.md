# standalone-diff Specification

## Purpose
TBD - created by archiving change standalone-diff-tool. Update Purpose after archive.
## Requirements
### Requirement: diff_two_files compares two files directly

`diff_two_files(old_path, new_path, output_dir)` SHALL 直接比较两个 xlsx 文件，不依赖 SVN。

#### Scenario: Two files compared successfully

- **WHEN** 传入两个已有 xlsx 文件路径和输出目录
- **THEN** 系统 SHALL 生成 diff Excel 到输出目录，并自动用 Excel 打开

### Requirement: Standalone diff CLI accepts two arguments

`diff_standalone.py` SHALL 接收两个命令行参数（旧文件、新文件），静默生成并打开 diff。

#### Scenario: TortoiseSVN invocation

- **WHEN** `diff.exe "%base" "%mine"` 被 TortoiseSVN 调用
- **THEN** 系统 SHALL 从 mine 路径推导项目根，输出到 `RAW_DIR/.temp/{stem}_diff.xlsx`，并静默打开

#### Scenario: Path detection fallback

- **WHEN** 从 mine 路径向上找不到 DataTable/Raw
- **THEN** 系统 SHALL fallback 到 `CATSINTHESHELL_ROOT` 环境变量，最后兜底 `A:\DataTable\Raw`

