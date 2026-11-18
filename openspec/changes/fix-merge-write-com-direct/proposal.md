# 合表工具写回链路修复：COM 直写替代 openpyxl 写回 + 全表重算

> 状态：已定稿（2026-08-06 副本 TableManager 实测验证通过），待实施。
> 本提案经历了完整的前因后果调查（见下「调查历程」），期间 `fix-external-link-target-recalc`（规范化 Target + 轮询重算）被提出并**证伪**，本提案为最终结论。

## Why（前因后果）

### 现象

`add-merge` / `merge resolve` 写回配置表（Effect.xlsx / Card.xlsx 等含外部链接的公式列表）后，大量外部引用公式缓存被重算成 `#N/A`（实测 Effect.xlsx #事件效果 2 → 2194 处），导表（DirectExcel / xlnt 读缓存值）出错。

### 调查历程（每步均有实测证据）

1. **初始假设 A：外部链接盘符不可达**。rels Target 为 `file:///A:\DataTable\Raw\...`，曾假设 A: 盘不存在 → 链接加载失败。
   - **证伪**：`A:` 是 `subst` 映射（`A:\: => D:\catsintheshell`），文件全部存在、`Workbooks.Open('A:\...')` 成功。

2. **初始假设 B：COM 自动化下链接数据永不加载**（交互式弹窗点「更新」才加载，`DisplayAlerts=False` 抑制弹窗 = 不更新）。
   - 部分成立但归因错误：新写入的**普通单元格引用**（`='[1]#Main'!A1`）确实永远 `#REF!`（普通引用需要外部工作簿真实打开）；但**配置表使用的 XLOOKUP 公式**不依赖此路径（见下）。

3. **探索方案 1：UpdateLink 显式传参 / ChangeLink / UpdateLinks 各参数 / AskToUpdateLinks / AutomationSecurity 组合**（含网络搜索到的方案，全部实测）。
   - **证伪**：`UpdateLink(LinkSources, 1)` 调用不抛错但不加载数据；`ChangeLink` 改 D: 真实路径后仍失败；**且 UpdateLink 会破坏链接数据**（调用后 XLOOKUP 外部引用全部 #N/A，test_unit_lookup 证据）。

4. **探索方案 2：规范化 rels Target（`file:///A:\` → `A:\` 无前缀 / 相对路径 / 正斜杠）+ 轮询等待链接加载后重算**（`fix-external-link-target-recalc`）。
   - 表面成功（H6 等公式"重算后正常"），**实为假阳性**：当时未经过 openpyxl 写回，缓存完好。
   - **决定性证伪实验**：对同一文件做「openpyxl 写回 + 外链/元数据注入恢复 + 规范化 + 轮询重算」，#N/A 667 → **2416**。规范化/轮询无法挽救 openpyxl 写回的破坏。

5. **根因锁定（openpyxl 写回三连击）**：openpyxl 保存对含外部链接公式列的表造成三重破坏：
   a. **清空全部公式缓存 `<v>`**（openpyxl 不保留缓存值）；
   b. **公式字符串中的中文字符被转义为 XML 实体**（`IF(I6="无",...)` → `IF(I6="&#26080;",...)`）——全量公式 diff 实测 7698 个公式被改写；
   c. **shared formula 全部展开为独立公式**（3586 → 10183 个公式节点）且**动态数组元数据（metadata.xml）丢失**（需 zip 注入恢复，注入后仍不完整）。
   叠加效果：重算时 XLOOKUP 的外部引用数据源失效 → 缓存被写成 #N/A。

6. **最终验证（副本 TableManager 实测）**：将写回改为 **COM 直写**（`UpdateLinks=0` 打开 → 写单元格 → `Save()`，不触发 `CalculateFull`）：
   - 新增行 / 修改静态单元格 / **修改名称列** 全部正确写入；
   - **Save 时 Excel 局部重算**（只算被改单元格的依赖链）**不会产生链接型 #N/A**——打开时（UpdateLinks=0）链接数据已加载，XLOOKUP 局部求值可用（实测 H9 Params 公式联动重算成功：`70030,2001100,31011016` → `70030,2001100,31011016,91011023`，公式保持公式）；
   - #N/A 基线 **667 → 667 零新增**；
   - 耗时 4.8s/表（含 Excel 冷启动，openpyxl 为 0.76s——可接受，可复用 Excel 实例/批量 Range 优化）。

### 关键认知（供后续维护者参考）

- **XLOOKUP（动态数组函数）跨工作簿引用依赖「打开时加载的链接数据」，与 rels Target 格式、UpdateLink 调用、externalLink XML 缓存均无直接关系**；普通单元格外部引用（`='[n]#Sheet'!A1`）则需要真实打开外部工作簿（本机 COM 下恒 #REF!，但配置表不使用此类引用）。
- **openpyxl 保存是破坏源头**（三连击），任何基于 openpyxl 写回 + Excel 重算的组合都无法修复。
- **全表重算（CalculateFullRebuild）禁止**：会重算所有外部引用（即使链接数据可用，#事件效果 等行存在空段/名称缺失的数据问题会成批暴露为 #N/A 并写回缓存）。
- 残留 #N/A（如 #事件效果 711 行「空段」：公式模板有 L~R 段但名称列留空 ≠「无」）是**表格数据问题**，由写回后校验暴露给策划修正，不属于本方案修复范围。

## What Changes

### 新增

- **`excel_recalc.write_cells_via_com(target_path, cells)`**：COM 直写入口（潜规则 2026-08-03 记载此函数但当前代码缺失，需补回）：
  `Workbooks.Open(UpdateLinks=0)` → 写静态值单元格（Cells(row, col).Value）→ `Save()` → 关闭。**不调用 CalculateFull**；写入顺序按行（避免依赖链错乱）。
- **写回后 #N/A 基线校验**（`table_checks` / `processed_checks`）：写回前记录各 sheet 的 `#N/A`/`#REF!` 缓存计数（基线），写回后对比；**新增错误值即报告**（区分链接问题与数据问题，数据问题列出涉及 sheet 供策划处理）。`count_missing_formula_caches` 扩展同时检查错误值。
- **公式写入拒绝**：staging/合并计划中若含公式单元格（`formula` 字段），COM 直写路径直接报错拒绝，提示改为静态值（公式写入会触发未受控求值）。

### 修改

- **`addmerge_engine._write_changes`**（410）：openpyxl 写回 → 改为 COM 直写（静态值优先）；openpyxl 写回仅保留为无 Excel 环境的降级路径（明确警告缓存将被清空、外链需注入恢复、文件可能无法被 Excel 打开）。
- **`addmerge_engine.apply_additive_merge`**（521-566）：`_recalc_with_degrade`（全表重算）删除；写回后接 #N/A 基线校验；事务备份/回滚链保持不变。
- **`merge_workflow` resolve/recover 链路**：同样改 COM 直写 + #N/A 校验；`_refresh_formula_cache`（238/263/315）移除。
- **`merge_engine._resolve_apply`**（1656-2159）：写回方式改 COM 直写。
- **`excel_recalc.ensure_formula_cache`**：仅保留给「无外部链接的表」的可选重算（有外部链接的表默认拒绝，避免误用）。

## 最小实施子集（MVP，已验证）

1. `write_cells_via_com` 落地（验证原型为临时副本中的 `_write_changes_com`，已随过程产物清理；按本提案方案实施，功能要点：UpdateLinks=0 打开 → Cells(row,col).Value 写值 → Save → Close，不触发 CalculateFull）；
2. add-merge 写回切换 + 删除全表重算 + 接入 #N/A 基线校验；
3. merge resolve 同步切换；
4. 回归：Card.xlsx（0 基线）与 Effect.xlsx（667 基线）真实文件副本全链路（新增行/静态覆盖/名称列修改），断言 #N/A 零新增、公式保持公式、文件可被 Excel 打开。

## Capabilities

### New Capabilities

- `com-direct-write`: 工具写回走 Excel COM 直写（UpdateLinks=0 + Save，无全表重算），外部引用缓存/公式/shared formula 原样保留
- `formula-cache-na-guard`: 写回前后 #N/A/#REF 基线对比校验，新增错误值报告（数据问题清单）
- `formula-write-reject`: staging/合并计划含公式单元格时拒绝（COM 直写不支持公式写入）

### Modified Capabilities

- `add-merge` / `merge-resolve`: 写回链路从「openpyxl + 全表重算」改为「COM 直写 + #N/A 校验」

## Impact

- `excel_recalc.py`：新增 `write_cells_via_com`；`ensure_formula_cache` 限定无外部链接表
- `addmerge_engine.py`：`_write_changes` (410)、`apply_additive_merge` (521-566) 改造
- `merge_workflow.py`：`_refresh_formula_cache` (189-195 / 238 / 263 / 315) 移除，写回改 COM
- `merge_engine.py`：`_resolve_apply` (1656-2159) 写回方式
- `table_checks.py` / `processed_checks.py`：`count_missing_formula_caches` 扩展错误值检查
- `tests/`：补 COM 直写、#N/A 基线校验、公式拒绝用例（`HAS_WIN32COM` 保护）

## 跨模块关注点

- COM 直写依赖本机 Excel（与现状一致）；无 Excel 环境降级 openpyxl 时**必须**警告缓存/外链风险并提示人工用交互式 Excel 打开→点「更新」→保存修复
- 写回单元格限定静态值；名称列修改安全（Save 局部重算可用），但写入值必须存在于对应查找域（否则 #N/A 由校验报告）
- 人工兜底：SVN 提交前交互式 Excel 打开 → 弹窗点「更新」→ 保存（刷新缓存，顺带清理存量 #N/A）
- 验证记录：2026-08-06 在临时副本 TableManager 上实测通过（新增行/静态覆盖/名称列修改全链路，#N/A 零新增、公式保持公式、耗时 4.8s/表）；测试环境已清理，实施时按本提案重建验证
