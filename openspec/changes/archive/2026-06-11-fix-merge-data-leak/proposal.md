## Why

`resolve_merge` 在插入 THEIRS_ADDED 行时，现有行的数据会泄漏到相邻行。TC15 测试中，mine 的 7001001 行（E="进攻性装备", F="1022"）的数据污染了 7000001（SAME 行，应为 E="", F=""）和 7001000（THEIRS_ADDED，应为 E="", F=""）。

根因在 `resolve` 的 Phase 2-4：`existing_rows` 按 mine 的 sheet row number 索引，但 `ordered` 列表中的 `out_pos` 在存在插入行时与实际写入位置产生偏移，导致 `existing_rows[old_row]` 的数据被写入错误的目标行。

同时，测试对比依赖 `data_only=True` 读取缓存值，但生成的输出由 openpyxl 写入（公式无缓存值），导致对比时大量公式单元格返回 None。需在对比前对输出文件做 Excel 重算。

## What Changes

### 1. fix resolve_merge `existing_rows` 写入位置映射

- **Phase 2 (ordered 列表构建)**: 修复 `out_pos` 与 `old_row` 的映射关系。当前 `out_pos` 按顺序自增，但 `insert_at` 是 sheet row number，两者在存在前面行的插入/删除时不一致。
- **Phase 4 (数据写入)**: 确保 `existing_rows[old_row]` 的全部列数据写入正确的 `write_row`，不混入 `final_row_data` 的无关列。

### 2. 测试流添加 Excel 重算

- `run_tests.py` Step 3: 在对比前对 resolved 输出和 correct 文件分别做 `excel_recalc.recalc_xlsx` 重算，确保公式缓存值存在。
- 恢复对比函数使用 `data_only=True`，移除公式文本 fallback（不再需要）。

### 3. run_tests.py 更新 TC15 文件名

- 用例名从 `TC15_Effect` 改为 `TC15_综合`
- 预期输出文件名从 `_resolved_correct.xlsx` 改为 `_correct.xlsx`

## Verification

修复后运行完整测试：

1. `python run_tests.py` — 14/14 PASS
2. TC15 专项验证：
   - 生成输出  vs `_correct.xlsx`：行数一致、行匹配 100%
   - Excel 重算后 value diff → 0
   - 数据泄漏消除：7000001 E/F 列为空，7001000 E/F 列为空

## Impact

- `merge_engine.py`: Phase 2/3/4 (`resolve` 方法内)
- `run_tests.py`: compare 流程 + TC15 配置
- 不影响 ConflictDetector、MergePreview
