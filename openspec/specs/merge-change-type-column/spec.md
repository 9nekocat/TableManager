# merge-change-type-column Specification

## Purpose
TBD - created by archiving change merge-preview-change-type-column. Update Purpose after archive.
## Requirements
### Requirement: 修改类型列显示
合并预览 Excel 的 B 列 SHALL 作为「修改类型」列。`[我的]` 和 `[仓库]` 行的 B 列 SHALL 显示对应侧的变更类型（`[已修改]` / `[已添加]` / `[已删除]`），仅作标注。

#### Scenario: 我的行标注已修改
- **WHEN** 某行在当前侧的变更为修改（MINE_MOD）
- **THEN** 对应的 `[我的]` 行 B 列显示 `[已修改]`

#### Scenario: 仓库行标注已添加
- **WHEN** 某行在当前侧的变更为新增（THEIRS_ADDED）
- **THEN** 对应的 `[仓库]` 行 B 列显示 `[已添加]`

#### Scenario: 我的行标注已删除
- **WHEN** 某行在我方被删除（MINE_DELETED）
- **THEN** `[我的]` 行 B 列显示 `[已删除]`

#### Scenario: 双方新增行的标注
- **WHEN** 某行为双方新增（BOTH_ADDED）
- **THEN** `[我的]` 行和 `[仓库]` 行 B 列均显示 `[已添加]`

### Requirement: 最终行下拉选择
`[最终]` 行的 B 列 SHALL 使用下拉列表（DataValidation），选项为 `[修改]` 和 `[删除]`。

#### Scenario: 下拉列表选项
- **WHEN** 用户点击 `[最终]` 行的 B 列单元格
- **THEN** 系统显示下拉选项 `[修改]` 和 `[删除]`

### Requirement: 最终行默认值
`[最终]` 行的修改类型列 SHALL 根据冲突类型设置默认值。

#### Scenario: 单方修改默认保留
- **WHEN** 冲突类型为 MINE_ADDED / THEIRS_ADDED / MINE_MOD / THEIRS_MOD
- **THEN** `[最终]` 行 B 列默认值为 `[修改]`

#### Scenario: 冲突类型默认保留
- **WHEN** 冲突类型为 CONFLICT 或 BOTH_ADDED
- **THEN** `[最终]` 行 B 列默认值为 `[修改]`

#### Scenario: 单方删除默认删除
- **WHEN** 冲突类型为 MINE_DELETED 或 THEIRS_DELETED
- **THEN** `[最终]` 行 B 列默认值为 `[删除]`

### Requirement: 决议器读取修改类型列
`MergeResolver` SHALL 读取 `[最终]` 行 B 列的值来决定该行的去留。

#### Scenario: 修改类型为修改
- **WHEN** B 列值为 `[修改]`（或为空）
- **THEN** 该行数据写入输出文件

#### Scenario: 修改类型为删除
- **WHEN** B 列值为 `[删除]`
- **THEN** 该行不写入输出文件（跳过）

### Requirement: 决议器不再依赖隐式逻辑
`MergeResolver` SHALL NOT 使用 `group_has_mine` + `base_order` 的隐式 3-way 分类来判定行去留。所有去留决定 SHALL 来自 `[最终]` 行 B 列的显式值。

#### Scenario: THEIRS_DELETED 被正确删除
- **WHEN** 仓库删除了某行且用户未修改 `[最终]` 行的默认值 `[删除]`
- **THEN** 决议器跳过该行，输出文件不含此行

