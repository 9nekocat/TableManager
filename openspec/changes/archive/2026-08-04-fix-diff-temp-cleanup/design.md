## Context

`diff_two_files()` 是 TortoiseSVN 外部 diff 工具入口（通过 `diff_standalone.py` 调用）。它比较两个 xlsx 文件，生成 diff Excel，并在 Excel 中打开。

临时文件创建在 `new_path.parent`（即工作副本所在 Raw 目录），以保证 Excel 能解析相对路径的外部链接引用。

## Goals / Non-Goals

**Goals:**
- 每次 diff 运行使用唯一临时文件名，避免跨次冲突
- 异常路径也清理临时文件

**Non-Goals:**
- 不改变临时文件的目录位置（必须在源文件同目录）

## Decisions

### Decision 1: 文件名加 UUID

`{stem}_{role}_recalc_{uuid}.xlsx` 替代固定名 `.{stem}_{role}_recalc.xlsx`

6 位 hex UUID 提供 16^6 = 16M 种可能，冲突概率极低。清理逻辑不变（finally 删除）。

**替代方案**: 用 `tempfile.mkstemp` → 文件名不含 stem，链接解析可能失败。不可用。

### Decision 2: 扩大 try/finally 范围

当前只有 `_read_all_rows_with_formulas` + `os.remove` 在 try/finally 中。`recalc_batch` 失败时临时文件未清理。

修复后整个 recalc + read 包裹在 try/finally 中。

## Risks

- [UUID 使文件名更长] → 不影响 Excel 或 openpyxl 的文件名长度限制
