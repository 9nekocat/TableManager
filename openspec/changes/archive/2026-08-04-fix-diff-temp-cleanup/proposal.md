## Why

`diff_two_files()` 在 Raw/ 目录创建固定路径临时文件 `.Effect_old_recalc.xlsx` / `.Effect_new_recalc.xlsx`。首次运行后 `os.remove` 可能因句柄未释放而静默失败，残留文件导致第二次比对异常退出，diff Excel 不弹出。

## What Changes

- `diff_two_files()`: 临时文件名加入 6 位随机 UUID → 每次运行路径唯一，避免跨次冲突
- `diff_two_files()`: `recalc_batch` + 读取整体包裹 try/finally → 异常路径也清理临时文件
- 批量重算临时文件保留在源文件同目录（保证链接工作簿的相对路径解析）

## Capabilities

### New Capabilities

- `diff-temp-cleanup`: diff 临时文件使用唯一命名 + 异常安全清理

## Impact

- `diff_engine.py`: `diff_two_files()` (lines 287-330)
