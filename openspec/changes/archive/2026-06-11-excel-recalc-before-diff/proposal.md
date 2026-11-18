## Why

openpyxl 保存文件时不计算公式，公式单元格缓存值过期（`#N/A`）。diff 读的是过期缓存值，导致大量误报差异。需要在 diff 前用 Excel COM 重算文件刷新缓存值。

## What Changes

- 新增 `excel_recalc.py`：封装 win32com Excel 重算逻辑
- 修改 `diff_engine.py`：diff/export 前对源文件执行重算
- 修改 `diff_standalone.py`：独立 diff 也执行重算

## Capabilities

### New Capabilities

- `excel-recalc`: 用 Excel COM 打开 xlsx → 重算所有公式 → 保存，刷新缓存值

### Modified Capabilities

- `diff-value-only`: diff 前自动重算，确保比较的是最新计算结果

## Impact

- `excel_recalc.py`：新文件
- `diff_engine.py`：`export_diff_to_excel`、`diff_two_files` 增加重算步骤
- `diff_standalone.py`：重算后 diff
