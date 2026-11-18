## Context

当前 diff 直接从源文件读 `data_only=True` 缓存值。openpyxl 保存后缓存值过期，导致 `#N/A` 等错误值被误报为差异。

## Goals / Non-Goals

**Goals:**
- diff 前用 Excel COM 重算文件，刷新缓存值
- 100% 兼容所有 Excel 函数（XLOOKUP、TEXTJOIN、跨文件引用）

**Non-Goals:**
- 不改 merge 保存逻辑
- 不在每次文件保存时重算（只在 diff 时）

## Decisions

**实现**: `excel_recalc.recalc_xlsx(filepath, timeout=60)` 
- `win32com.client.Dispatch("Excel.Application")` 打开文件
- `Application.CalculateFull` 强制重算
- 关闭前自动保存

**调用点**: 
- `diff_engine.py:export_diff_to_excel` → 重算当前文件
- `diff_engine.py:diff_two_files` → 重算新文件（旧版 SVN 导出的临时文件已过期，但 SVN 版本可能是 Excel 保存的，可跳过）

**重算策略**: 先 copy 源文件到 TEMP_DIR 的副本，重算副本，读副本做 diff。不修改原始文件。

## Risks / Trade-offs

- Excel COM 启动开销 ~2-3 秒，diff 操作变慢但可接受
- 如果 Excel 已在运行，复用进程更快
- 跨文件 XLOOKUP 依赖外部 xlsx 存在 → 重算时会尝试更新链接
