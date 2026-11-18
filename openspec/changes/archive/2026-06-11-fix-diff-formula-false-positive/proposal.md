## Why

Diff 工具在比对两个版本时，会将仅公式变化而计算值未变化的单元格误报为差异。根因是源文件中部分单元格的缓存值（`data_only=True` 读取结果）变成了 `format_display` 格式化字符串（如 `=_xlfn.TEXTJOIN(...) → VALUE`），而非纯数值。比较时公式部分不同但值相同，却仍被判定为差异。

## What Changes

- `_cell_compare_str`：从 `val` 中提取纯值部分再比较——若 `val` 包含 `→` 分隔符，只取箭头后的值
- `_cell_display`：同样逻辑，确保 diff 显示的是值而非公式文本

## Capabilities

### New Capabilities

- `diff-value-only`: Diff 比对仅基于计算结果值，忽略公式文本差异

### Modified Capabilities

<!-- No existing capabilities to modify -->

## Impact

- `diff_engine.py`：`_cell_compare_str`、`_cell_display` 两个函数
- 不影响 merge 逻辑、不影响 SVN 操作
