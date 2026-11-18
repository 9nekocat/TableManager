# diff-value-only Specification

## Purpose
TBD - created by archiving change fix-diff-formula-false-positive. Update Purpose after archive.
## Requirements
### Requirement: Diff comparison uses computed values only

Diff 比对在判断两个单元格是否相同时，SHALL 仅基于计算结果值，忽略公式文本差异。

#### Scenario: Formula changed, value unchanged

- **WHEN** 旧版单元格缓存值为 `=_xlfn.TEXTJOIN(...,[5]...,K7) → 0,3001,5.1`，新版缓存值为 `0,3001,5.1`
- **THEN** 系统 SHALL 判定该单元格无差异

#### Scenario: Pure cached value comparison

- **WHEN** 旧版和新版均有纯数值缓存值，且两值相同
- **THEN** 系统 SHALL 判定该单元格无差异

#### Scenario: Value actually changed

- **WHEN** 旧版值为 `100`，新版值为 `200`
- **THEN** 系统 SHALL 判定该单元格有差异

### Requirement: Diff display shows computed values only

Diff 结果展示时，SHALL 仅显示计算结果值，不显示公式文本。

#### Scenario: Cell with format_display value

- **WHEN** 单元格缓存值为 `=SUM(A1:A5) → 150`
- **THEN** diff 展示 SHALL 显示 `150`

#### Scenario: Cell with clean value

- **WHEN** 单元格缓存值为 `200`
- **THEN** diff 展示 SHALL 显示 `200`

#### Scenario: Empty cell

- **WHEN** 单元格缓存值为 `None`
- **THEN** diff 展示 SHALL 显示空字符串

