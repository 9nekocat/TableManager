## 1. Excel Recalc

- [x] 1.1 创建 `excel_recalc.py`，实现 `recalc_xlsx(filepath, timeout=60)`
- [x] 1.2 处理 Excel 进程生命周期（复用已有实例、超时强制关闭）

## 2. Diff Integration

- [x] 2.1 修改 `diff_engine.py:export_diff_to_excel` — diff 前 copy 当前文件到 TEMP_DIR 并重算
- [x] 2.2 修改 `diff_engine.py:diff_two_files` — diff 前重算 `new_path`

## 3. Build & Verify

- [x] 3.1 用 Card.xlsx 验证：diff r4583 时 col36/38/40 不再显示 `#N/A` 误报
- [x] 3.2 重新打包 `TableManager.exe` 和 `diff.exe`
