## 1. Core Fix

- [x] 1.1 在 `diff_engine.py` 新增 `_extract_value(cell_tuple)` 函数：从 `(val, fml)` 元组提取纯值，若 `val` 含 `→` 则取箭头后部分
- [x] 1.2 修改 `_cell_compare_str`：使用 `_extract_value` 替代直接取 `val`
- [x] 1.3 修改 `_cell_display`：使用 `_extract_value` 替代直接取 `val`

## 2. Unit Tests

- [x] 2.1 新建 `tests/test_diff_engine.py`，覆盖 `_extract_value`、`_cell_compare_str`、`_cell_display`
- [x] 2.2 测试用例：纯数值、format_display 格式（含 `→`）、None、空字符串、公式变值不变、值真变
- [x] 2.3 运行测试确保全部通过

## 3. Verify

- [x] 3.1 用现有 `Effect_diff_r4583.xlsx` 同类数据验证：公式变化值不变时不再误报差异
- [x] 3.2 确认值实际变化时仍能正常检测差异
