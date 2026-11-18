## Context

`_cell_compare_str` 和 `_cell_display` 从 `(val, fml)` 元组中取 `val` 进行比较和显示。`val` 来自 `_read_all_rows_with_formulas` 中 `data_only=True` 加载的缓存值。

当前问题：部分源文件（尤其是通过本工具的 merge 流程处理的文件）的缓存值变成了 `format_display` 格式化字符串（`=FORMULA → VALUE`），而非纯数值。导致 `_cell_compare_str` 比较的是包含公式文本的完整字符串，公式差异（如 `[5]` vs `[4]`）被误判为值差异。

## Goals / Non-Goals

**Goals:**
- `_cell_compare_str` 只比较计算结果值，忽略公式文本差异
- `_cell_display` 只显示值，不显示公式文本

**Non-Goals:**
- 不修复源文件中缓存值异常的问题（那是另一个变更）
- 不涉及 merge 逻辑
- 不修改 `_read_all_rows_with_formulas`

## Decisions

**方案：在 `_cell_display` / `_cell_compare_str` 中做防御性提取**

两个函数都调用同一个 helper `_extract_value(cell_tuple)`：
1. 若 `cell_tuple` 为 None → 返回 `""`
2. 取 `val = cell_tuple[0]`
3. 若 `val` 是字符串且含 `→`（`format_display` 箭头），拆分取箭头后的值部分
4. 否则返回原值

选择此方案理由：
- 最小修改面，仅改两函数
- 向后兼容纯数值情况
- 防御未来可能的 format_display 泄漏

替代方案（不采用）：
- 直接修复 `_read_all_rows_with_formulas` 使其不返回 format_display → 需改上游逻辑，影响 merge 流程
- 用公式引擎求值 → 复杂度高，需引入外部依赖

## Risks / Trade-offs

- 如果值本身包含 `→` 字符（极罕见），会被错误截断。概率极低，且即使是 UE DataTable 的中文描述也不会用此字符
