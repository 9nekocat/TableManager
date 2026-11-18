# Merge Preview Display

合并预览生成的 `[我的]`/`[仓库]`/`[最终]` 三行布局的正确数据显示。

## ADDED Requirements

### Requirement: Added rows display correct side data

对于新增行（MINE_ADDED, THEIRS_ADDED, BOTH_ADDED），系统 SHALL 在 `[我的]`/`[仓库]` 行显示对应 side 的实际数据，而非空白或 base 数据。

#### Scenario: Mine-added row shows mine data in preview
- **WHEN** 生成 MINE_ADDED 行的合并预览
- **THEN** `[我的]` 行也显示 mine 侧数据，`[仓库]` 行不生成（该侧无标签）

#### Scenario: Theirs-added row shows theirs data in preview
- **WHEN** 生成 THEIRS_ADDED 行的合并预览
- **THEN** `[我的]` 行不生成（该侧无标签），`[仓库]` 行显示 theirs 侧数据

#### Scenario: Both-added row shows both sides
- **WHEN** 生成 BOTH_ADDED 行（双方都新增但数据不同）的合并预览
- **THEN** `[我的]` 行显示 mine 数据，`[仓库]` 行显示 theirs 数据，冲突单元格红色高亮

### Requirement: Deleted rows show base data for review

对于删除行（MINE_DELETED, THEIRS_DELETED），系统 SHALL 在对应 label 行显示 base 版本数据，让用户看到"被删除的内容"。

#### Scenario: Mine-deleted row shows base data
- **WHEN** 生成 MINE_DELETED 行的合并预览
- **THEN** `[我的]` 行标注 `[已删除]` 并显示 base 数据，`[仓库]` 行不生成

#### Scenario: Theirs-deleted row shows base data  
- **WHEN** 生成 THEIRS_DELETED 行的合并预览
- **THEN** `[我的]` 行不生成，`[仓库]` 行标注 `[已删除]` 并显示 base 数据

### Requirement: Non-conflict rows default to modifying side

对于非冲突行（单方修改/新增/删除），`[最终]` 行 SHALL 默认预填修改方的数据及对应决策。

#### Scenario: MINE_MOD row defaults to mine data with [修改]
- **WHEN** 生成 MINE_MOD 行的合并预览
- **THEN** `[最终]` 行 B 列为 `[修改]`，数据列预填 mine 值

#### Scenario: THEIRS_MOD row defaults to theirs data with [修改]
- **WHEN** 生成 THEIRS_MOD 行的合并预览
- **THEN** `[最终]` 行 B 列为 `[修改]`，数据列预填 theirs 值

#### Scenario: MINE_DELETED row defaults to [删除]
- **WHEN** 生成 MINE_DELETED 行的合并预览
- **THEN** `[最终]` 行 B 列为 `[删除]`，数据列预填 base 值

#### Scenario: THEIRS_DELETED row defaults to [删除]
- **WHEN** 生成 THEIRS_DELETED 行的合并预览
- **THEN** `[最终]` 行 B 列为 `[删除]`，数据列预填 base 值

### Requirement: Conflict rows highlight differences

对于冲突行，`[最终]` 行的冲突单元格 SHALL 红色高亮，提醒用户需要手动决策。

#### Scenario: Conflict row has red-highlighted cells
- **WHEN** 生成 CONFLICT 行的合并预览
- **THEN** `[最终]` 行中冲突单元格使用浅红底色（FFCCCC），非冲突单元格使用绿色底色
