## Purpose

定义增量合并（add-merge，staging → Raw 目标表）写回路径对公式写入的拒绝规则：staging / 合并计划中若含公式单元格（`formula` 字段），直接报错拒绝并提示改为静态值，避免公式写入触发未受控求值造成缓存损坏。merge resolve 场景的决策行公式不适用本规则（由 com-direct-write 的受控公式写入覆盖）。

## ADDED Requirements

### Requirement: add-merge 的公式单元格拒绝写回

增量合并（add-merge）写回路径 SHALL 拒绝公式写入：写回计划中含 `formula` 字段的单元格 SHALL 触发错误，错误信息 SHALL 提示改用静态值。拒绝发生时目标文件 SHALL 保持未被修改（事务回滚）。

#### Scenario: 写回计划含公式单元格被拒绝

- **WHEN** add-merge 写回计划包含公式单元格（如 `=B4+1`）
- **THEN** 写回 SHALL 抛错并提示「公式写入不支持，请改为静态值」，且目标文件不被写入

#### Scenario: 纯静态值计划正常写回

- **WHEN** add-merge 写回计划全部为静态值单元格
- **THEN** 写回 SHALL 正常执行，不受公式拒绝规则影响
