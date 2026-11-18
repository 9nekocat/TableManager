## Why

`generate_diff("Effect.xlsx", "HEAD")` 耗时 10s，其中 6.5s 消耗在 Excel COM 重算（旧版本 4.4s + 当前版本 2.1s）。SVN 导出的旧版本和 Excel 保存的工作副本通常已有缓存公式值，recalc 是冗余操作。另外 `diff_two_files` 逐个调用 `recalc_xlsx` 打开两次 Excel 实例（浪费 COM 启动开销）。

## What Changes

- `diff_engine.py` `_get_old_version()`: 去掉 copy-to-Raw + `recalc_xlsx`，直接从 SVN temp 文件读取
- `diff_engine.py` `generate_diff()` / `export_diff_to_excel()`: 去掉 `recalc_to_temp`，直接读工作副本
- `diff_engine.py` `diff_two_files()`: 用 copy-to-temp + `recalc_batch`（一次 Excel 会话处理两个文件），确保公式缓存值可用
- `excel_recalc.py`: 新增 `recalc_batch()` 函数 — 一次 Excel COM 会话打开多文件 + 链接工作簿 + 统一重算
- `diff_engine.py` `_cell_display()`: 恢复到纯值显示（公式缓存值由 recalc 保证）

## Capabilities

### New Capabilities

- `batch-excel-recalc`: 批量 Excel 重算——在一次 Excel COM 会话中重算多个 xlsx 文件的公式缓存值
- `diff-no-recalc-optimized`: 在不需重算的场景（普通工作副本 vs SVN 版本）跳过 Excel recalc，直接读取

## Impact

- `diff_engine.py`: `_get_old_version()`, `generate_diff()`, `export_diff_to_excel()`, `diff_two_files()`, `_cell_display()`
- `excel_recalc.py`: 新增 `recalc_batch()`
- `diff_standalone.py`: 通过 `diff_two_files()` 间接受益
