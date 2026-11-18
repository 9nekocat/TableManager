# com-direct-write Specification

## Purpose
定义配置表工具写回链路的统一行为契约：通过 Excel COM 直写（UpdateLinks=0 打开 + 写静态值 + Save，不触发全表重算）替代 openpyxl 写回，使外部引用公式缓存、公式字符串与 shared formula 结构在写回后原样保留，不再产生 #N/A 缓存破坏。
## Requirements
### Requirement: 写回使用 COM 直写

工具写回入口 SHALL 通过 Excel COM 完成：`Workbooks.Open(UpdateLinks=0)` 打开目标工作簿 → 按行序写入静态值单元格（`Cells(row, col).Value`）→ `Save()` → 关闭。写回全程 SHALL NOT 调用 `CalculateFull` 或任何等价的全表重算。

#### Scenario: 含外部链接公式的表写回后缓存保留

- **WHEN** 对包含外部链接 XLOOKUP 公式列（如 Effect.xlsx #事件效果）的配置表执行写回
- **THEN** 写回后外部引用公式缓存 SHALL 与写回前一致，不新增 #N/A 错误值

#### Scenario: 修改被公式引用的单元格后局部重算

- **WHEN** 修改名称列等被公式引用的静态单元格并 `Save()`
- **THEN** Excel SHALL 仅对受影响单元格的依赖链做局部重算，公式单元格保持公式且缓存值正确联动

#### Scenario: 写入按行序执行

- **WHEN** 写回计划包含多行多单元格写入
- **THEN** SHALL 按行序写入（同列跨行、同列引用依赖正确），避免依赖链错乱

### Requirement: 写回不触碰未变更内容

写回 SHALL 只写入变更单元格，未变更的公式、缓存值、外部链接与动态数组元数据 SHALL 原样保留，无需 zip 层注入恢复。

#### Scenario: 未变更公式保持原样

- **WHEN** 表中某公式单元格不在写回计划内
- **THEN** 该单元格的公式字符串、缓存值与 shared formula 结构 SHALL 与写回前完全一致

### Requirement: 无 Excel 环境的降级路径

win32com/Excel 不可用时，写回 SHALL 降级为 openpyxl 全量写回，且 MUST 明确警告：公式缓存将被清空、外部链接与动态数组元数据需要 zip 注入恢复、文件可能无法被 Excel 正常打开。

#### Scenario: 无 Excel 时降级并警告

- **WHEN** 本机无 win32com 且执行写回
- **THEN** 写回 SHALL 以 openpyxl 完成，并在返回值/报告中明确标记降级风险

