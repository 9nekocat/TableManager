# 设计：COM 直写写回链路

> 状态：基于已定稿提案（2026-08-06 副本实测验证通过）的实施设计。
> 范围：add-merge / merge resolve / recover 三条写回链路从「openpyxl 写回 + 全表重算」切换为「COM 直写 + #N/A 基线校验」。
> 注：回归目标真实文件为 `CardTable.xlsx`（提案中写 Card.xlsx）与 `Effect.xlsx`。

## 1. excel_recalc.py

### 1.1 新增 `write_cells_via_com(target_path, cells)`

COM 直写入口（add-merge 使用），语义（提案「写回单元格限定静态值」）：

```python
def write_cells_via_com(target_path, cells, timeout=120):
    """COM 直写静态值到目标 xlsx：UpdateLinks=0 打开 → 按行序写值 → Save → 关闭。
    cells: {sheet_name: [(row, col, value), ...]}；value 必须为静态值。
    含公式字符串（str 以 '=' 开头）→ 抛 ValueError（公式写入拒绝）。
    """
```

- 流程：`pythoncom.CoInitialize()` → `DispatchEx("Excel.Application")`（私有实例，禁止 `Dispatch()` 附着用户实例，与 `recalc_xlsx` 一致）→ `_configure_excel()` → `Workbooks.Open(fullpath, UpdateLinks=0)` → 逐 sheet 按行序 `ws.Cells(row, col).Value = value`（行序写入避免依赖链错乱）→ `Workbook.Save()` → `Close(SaveChanges=False)` → `excel.Quit()` → `CoUninitialize()`。
- **不调用 `CalculateFull`**：只依赖 Save 时 Excel 的局部重算（打开时链接数据已加载，XLOOKUP 局部求值可用，实测 H9 Params 联动成功）。
- **必须覆盖 `CalculateBeforeSave = False`**：`_configure_excel` 默认设 `True`，会让 Save 前重算整个工作簿，把「空段/名称缺失」等存量数据问题成批暴露为**新的** #N/A 缓存（实测 Effect #事件效果 2 → 235，全表重算行为）。COM 直写两个入口都在 `_configure_excel` 后显式覆盖为 `False`；`recalc_xlsx`（重算场景）保留 `True`。
- 公式值拒绝：`isinstance(value, str) and value.startswith("=")` → `raise ValueError("COM 直写不支持公式写入，请改为静态值: sheet row col")`。
- 错误清理：`finally` 中依次关闭 workbook / excel / CoUninitialize，防止 Windows 句柄泄漏（对齐 `recalc_xlsx` 的清理结构）。
- 无 win32com：抛 `RuntimeError`，由调用方决定降级。

### 1.2 `count_missing_formula_caches` 扩展错误值检查

- 新增 `count_formula_error_caches(filepath)` → `{sheet_name: count}`：沿用 XML 层解析（`_sheet_xml_targets` + iterparse，避免 openpyxl data_only 的 None/空串歧义），统计缓存值为错误值（`t="e"` 或 `<v>` 文本以 `#` 开头，如 `#N/A`、`#REF!`）的公式单元格，按 `#` sheet 分组。
- `count_missing_formula_caches` 保持现有语义不变（避免影响 diff/merge 现有调用），错误值统计走独立函数，二者可叠加使用。
- table_checks / processed_checks 如需报告错误值，复用 `count_formula_error_caches`（当前这两个模块未直接调用 `count_missing_formula_caches`，无强制改动，实施时确认）。

### 1.3 新增 #N/A 基线校验助手

```python
def verify_formula_cache_guard(filepath, baseline, data_start_row=4):
    """写回后错误值基线校验：与写回前 baseline 对比。
    返回 {ok, baseline, after: {sheet: count}, added_sheets: [...], total}
    基线已存在错误值视为表格数据问题，不算失败；新增错误值 ok=False。
    """
```

- 基线采集 = 写回前 `count_formula_error_caches()`（在事务备份之后、写回之前执行，保证基准是未修改文件）。
- 新增错误值列表（sheet 级）进入报告，供策划定位数据问题；写入 `report["na_guard"]`（字段对齐 formula-cache-na-guard spec）。

### 1.4 `ensure_formula_cache` 限定无外部链接表

- 入口处检查 `_read_external_link_cache(filepath)`：含外部链接 → 拒绝执行全表重算，`raise_on_error=True` 时抛 `RuntimeError`（提示写回已改 COM 直写，勿手动全表重算），否则返回 `{"ok": False, "error": ...}`。
- 保留给无外部链接表的可选重算（diff 等场景仍可用）；有外部链接的表默认拒绝，避免误用（提案：CalculateFull 会重算所有外部引用并成批写 #N/A 缓存）。

## 2. addmerge_engine.py

### 2.1 `_write_changes`（410）改 COM 直写

- 保留 openpyxl 读取目标表（data_only=False）仅用于**计算写入计划**：modified 行的目标坐标（`_row_of`）、added 行的插入位置（`_last_content_row` + 递增）。计算结果转成 `{sheet: [(row, col, value)]}`。
- **公式拒绝**：`ch["formula"]` / `w["formula"]` 非空 → `raise AddMergeError("staging 含公式单元格，COM 直写不支持公式写入，请改为静态值: sheet row col")`（对齐 formula-write-reject spec）。原 `adjust_formula` / `ArrayFormula` 写回逻辑随 openpyxl 路径迁移至降级分支。
- 写回方式：win32com 可用 → `write_cells_via_com(target_path, writes)`；不可用 → 降级 openpyxl 写回 + 外链注入 + 元数据注入（现有逻辑保留），并在返回值/日志明确警告「缓存将被清空、外链需注入恢复、文件可能无法被 Excel 打开」。
- **COM 成功路径不再需要 `_inject_external_links` / `_inject_cached_metadata`**（COM Save 天然保留外链与动态数组元数据），`extlink_cache` / `metadata_xml` 参数仅降级分支使用。

### 2.2 `apply_additive_merge`（521-566）

- 事务链：备份 → **写回前基线采集** → `_write_changes` → `_verify_external_links_preserved` → `verify_formula_cache_guard`。
- 删除 `_recalc_with_degrade`（507）及其调用；`recalc_timeout` 参数保留（签名兼容，供降级分支透传）。
- `report["formula_cache"]` → `report["na_guard"]`（baseline / after / added_sheets / ok）；无 Excel 降级时 `na_guard = {"ok": None, "skipped": "..."}`。
- 备份/回滚链不变（失败时 `restore_merge_backup` + `AddMergeError(backup_path=, restored=)`）。

## 3. merge_workflow.py

- 删除 `_refresh_formula_cache`（189-195）。
- `apply_merge_preview_transactional`（231-307）与 `recover_merge_preview_transactional`（308-342）：
  - 移除 `recalc_func` / `recalc_timeout` 参数与调用（提案 Impact 的 238/263/315）；`resolve_merge` 后接 `verify_formula_cache_guard`（基线与 add-merge 相同：备份后、resolve 前采集）。
  - 结果 dict `"formula_cache"` → `"na_guard"`；`normalize_external_link_targets_to_preferred_drive` 与 `_verify_external_links_preserved` 保留。
  - `recover_merge_preview_transactional` 透传参数同步删除。
- 调用方更新：`table_server.py:538/580`、`table_manager.py:473/523` 去掉 `recalc_func`/`recalc_timeout` 实参；`merge_workflow.py:330` 内部调用同步。

## 4. merge_engine.py `_resolve_apply`（1656-2159）

- 内存中 openpyxl 合并计算**原样保留**（决定合并结果：冲突决策、行插入位置、尾部行删除、公式调整），仅替换落盘阶段。
- 计算阶段改为**同时收集写计划**（`com_plan[sn]`）：
  - `created` / `headers`：theirs-only sheet 及表头；
  - `inserts`: `[(out_row, {col: value})]` — 新增行（out_row 为**最终行号**）；
  - `updates`: `[(row, col, value)]` — 既有行被 [最终] 决策修改的单元格（与 mine 原值不同才收集；决策 = mine 公式的单元格不收集——Excel 插入/删行时自动调整引用）；
  - `deletes` / `delete_from`：中间删除（mine 行号）与尾部删除。
- **公式受控写入**（resolve 决策行公式必须保留——Card.xlsx TEXTJOIN 数组公式回归）：公式值不拒绝，由 `apply_write_plan_via_com` 区分写入方式：str 以 `=` 开头 → `.Formula`；ArrayFormula 对象 → `.FormulaArray`（保留 CSE 数组语义）。UpdateLinks=0 打开时链接数据已加载，写公式后 Excel 局部求值可用（proposal 验证「公式保持公式」）。
- 落盘：win32com 可用 → COM 应用写计划（打开 `UpdateLinks=0` → 中间删行（降序）→ 尾部删除 Rows.Delete → 行插入 Rows.Insert（**按最终行号升序**，上方插入整体下推下方内容，先插上方行号才稳定；Excel 自动调整公式引用）→ 按行序写值 → Save）；不可用 → 保留现有 `out_wb.save` + `_inject_external_links` + `_inject_dynamic_array_metadata` 降级路径并警告。
- 注意：COM 直写必须在 **mine 原文件**上应用（非 out_wb），由 merge_workflow 的备份保证失败可回滚。

## 5. 测试（tests/）

- `test_na_guard.py`（无 Excel 可跑）：`count_formula_error_caches` 对构造含 `t="e"` / `#N/A` 缓存的小 xlsx（openpyxl 构造 + zip 层改 XML）；`verify_formula_cache_guard` 基线对比；公式拒绝路径（`write_cells_via_com` 无 win32com 时抛错分支）。
- `test_write_cells_com.py`（`HAS_WIN32COM` 保护，skip 条件对齐 `test_addmerge_engine.py:26-31`）：副本上新增行/静态覆盖/名称列修改，断言 #N/A 零新增、公式保持公式、openpyxl 可重新打开。
- 既有用例回归：`test_addmerge_engine` / `test_merge_workflow_backup` / `test_merge_conflict_decision`（确认 `formula_cache` 字段变化与签名移除不破坏）。

## 6. 回归验证（本机 Excel，人工）

- 数据：`DataTable/Raw/CardTable.xlsx`（0 基线）与 `Effect.xlsx`（667 基线）的**临时副本**（TEMP_DIR），绝不动真表。
- 全链路：add-merge 新增行 / 静态覆盖 / 名称列修改；merge resolve 冲突决策应用。
- 断言：#N/A/#REF 零新增、公式保持公式、`#事件效果` 行号不漂移、文件可被 Excel 打开、耗时记录（预期 ~5s/表）。

## 7. 风险与待验证

- **公式受控写入的求值时机**：`.Formula` / `.FormulaArray` 写入会触发该单元格的局部求值（链接数据已加载，proposal 已验证安全）；若公式引用本表未重算的依赖链，求值结果可能基于旧缓存——回归以 #N/A 零新增断言验证。
- **数组公式 FormulaArray 写入**：Excel COM 的 `.FormulaArray` 保留 CSE 语义，但写入的公式必须是区域无关的单单元格数组公式（如 TEXTJOIN 自引用），回归验证 Effect/Card 副本。
- **行插入后公式引用**：COM Rows.Insert 由 Excel 原生调整下方公式引用，与 openpyxl 的 `adjust_formula` 结果应一致，回归中以行号不漂移断言验证。
- **`formula_cache` 字段改名**影响 GUI/CLI 消费方：`table_server.py` / `table_manager.py` 展示逻辑确认 `na_guard` 渲染。
