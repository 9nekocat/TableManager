# Tasks

## excel_recalc.py

- [x] 新增 `write_cells_via_com(target_path, cells)`：COM 直写入口（CoInitialize → DispatchEx 私有实例 → UpdateLinks=0 打开 → 按行序 `Cells(row, col).Value` → Save → 清理）；静态值限定，公式字符串抛 ValueError；win32com 缺失抛 RuntimeError
- [x] 新增 `count_formula_error_caches(filepath)`：XML 层解析统计 `#` sheet 的错误值缓存（`t="e"` 或 `<v>` 以 `#` 开头），返回 `{sheet: count}`
- [x] 新增 `verify_formula_cache_guard(filepath, baseline)`：写回后错误值基线对比，返回 `{ok, baseline, after, added_sheets, total}`
- [x] `ensure_formula_cache` 增加外部链接守卫：文件含外部链接（`_read_external_link_cache` 非空）时拒绝全表重算

## addmerge_engine.py

- [x] `_write_changes` 改 COM 直写：openpyxl 仅用于计算写入计划；`formula` 字段 → AddMergeError 拒绝；win32com 可用走 `write_cells_via_com`，不可用降级 openpyxl + 外链/元数据注入并明确警告；COM 路径不再注入
- [x] `apply_additive_merge`：删除 `_recalc_with_degrade` 调用；写回前采集错误值基线；写回后 `verify_formula_cache_guard`；`report["formula_cache"]` → `report["na_guard"]`；备份/回滚链不变

## merge_workflow.py

- [x] 删除 `_refresh_formula_cache`；`apply_merge_preview_transactional` / `recover_merge_preview_transactional` 移除 `recalc_func` / `recalc_timeout` 参数与调用，resolve 后接 #N/A 基线校验，结果 `"formula_cache"` → `"na_guard"`
- [x] 更新调用方：`table_server.py:538/580`、`table_manager.py:473/523`、`merge_workflow.py:330` 去掉 recalc 实参

## merge_engine.py

- [x] `_resolve_apply` 落盘改 COM 直写：内存合并计算保留，同时收集写计划（行插入/单元格修改/尾部删除）；公式值拒绝；win32com 可用走 COM（删尾行 → 插行 → 按行序写值 → Save），不可用保留 openpyxl save + 注入降级并警告

## 测试与回归

- [x] 新增 `tests/test_na_guard.py`：错误值计数、基线对比、公式拒绝（无 Excel 可跑）
- [x] 新增 `tests/test_write_cells_com.py`（`HAS_WIN32COM` 保护）：副本全链路，断言 #N/A 零新增、公式保持公式
- [x] 既有 tests 全量通过（130 passed，含适配后的数组公式/外链/回滚用例）
- [x] 回归验证（本机 Excel）：`Event.xlsx`（0 基线）与 `Effect.xlsx`（667 基线）临时副本全链路（新增行/静态覆盖/冲突决策），断言 #N/A 零新增（Event 0→0、Effect 667→667，存量 #N/A 被 COM Save 顺带修复）、公式保持公式（A4 `=B4`）、公式 staging 拒绝+回滚、文件可被 Excel 打开；既有 tests 全量通过（130 passed）
