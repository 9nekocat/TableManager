## Why

合并预览当前用的是隐式逻辑表达变更类型：行颜色和 `[最终]` 行的出现/缺失承载了太多隐含义。THEIRS_DELETED 默认保留（违反"单方修改自动接受"原则），且用户无法在预览中明确表达"保留还是删除"的决策。

## What Changes

- 新增「修改类型」列（[标记] 列右侧），标注每行的变更性质
- `[我的]` / `[仓库]` 行在该列显示：`[已修改]` / `[已添加]` / `[已删除]`
- `[最终]` 行在该列使用下拉选择：`[修改]` 或 `[删除]`
- THEIRS_DELETED：默认决策从「保留」改为「删除」（在 [最终] 行预填 `[删除]`）
- MINE_DELETED：不再隐藏 [最终] 行，改为显式展示并预填 `[删除]`
- 决议器改为读取 [最终] 行的修改类型列决定行去留
- **BREAKING**：合并预览 Excel 列布局变化（新增一列），旧预览文件不兼容新决议器

## Capabilities

### New Capabilities

- `merge-change-type-column`: 合并预览新增「修改类型」列，在 `[最终]` 行通过下拉选择 `[修改]` 或 `[删除]` 来决定行的去留

### Modified Capabilities

<!-- No existing specs to modify -->

## Impact

- `merge_engine.py`：`MergePreview.generate`（列布局、下拉验证）、`MergeResolver.resolve`（读取新列、删除逻辑）
- 所有使用合并预览的工作流均受影响
