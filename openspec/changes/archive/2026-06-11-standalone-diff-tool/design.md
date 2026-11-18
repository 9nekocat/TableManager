## Context

当前 `export_diff_to_excel` 依赖 SVN 拉取旧版本。TortoiseSVN 外部 diff 工具方式是传入两个已导出的文件路径，需直接比较。

## Goals / Non-Goals

**Goals:**
- 接收两个绝对路径参数，生成 diff Excel 到 `.temp/`
- 静默打开 Excel（无控制台输出）
- 轻量打包（不含 Flask）

**Non-Goals:**
- 不改动现有 Web diff 功能
- 不添加 GUI

## Decisions

**`diff_two_files(old_path, new_path, output_dir)`** — 接收绝对路径，复用现有 `_read_all_rows_with_formulas` + `_compare_data` + Excel 写逻辑。

**路径推导**: `diff_standalone.py` 在 import diff_engine 前从 mine 路径推导项目根：
1. 从 mine 文件向上找 `DataTable/Raw`
2. 设置 `CATSINTHESHELL_ROOT` 环境变量
3. 兜底用 `A:\DataTable\Raw`

**PyInstaller spec**: 只打包必要依赖（diff_engine, merge_engine, formula_utils, manager_core），不含 Flask + templates。预估体积 ~15MB。

## Risks / Trade-offs

- merge_engine 的 `_read_all_rows_with_formulas` 被 diff 复用，但 merge_engine 导入了 manager_core 的 RAW_DIR → 依赖 `_find_project_root`。通过设置环境变量绕过。
