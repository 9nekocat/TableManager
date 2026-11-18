## Purpose

定义增量合并（add-merge，staging → Raw 目标表）写回链路的行为契约：写回改走 COM 直写保留公式缓存，写回后执行错误值基线校验，无 Excel 环境降级 openpyxl 并警告风险，事务备份/回滚链保持不变。

## ADDED Requirements

### Requirement: 写回走 COM 直写

增量合并写回 SHALL 使用 COM 直写（静态值优先）；openpyxl 写回仅保留为无 Excel 环境的降级路径，且 MUST 明确警告缓存将被清空、外链/元数据需注入恢复、文件可能无法被 Excel 正常打开。

#### Scenario: 有 Excel 环境写回保留缓存

- **WHEN** 本机存在 win32com/Excel 且 `apply=True`
- **THEN** 写回 SHALL 通过 COM 直写完成，外部引用缓存不新增 #N/A，公式保持公式

#### Scenario: 无 Excel 环境降级

- **WHEN** 本机无 win32com/Excel
- **THEN** 写回 SHALL 降级 openpyxl 路径并在报告中明确标记降级风险

### Requirement: 写回后错误值基线校验

增量合并写回成功后 SHALL 执行错误值基线校验，校验结果（baseline / after / added_sheets / ok）SHALL 进入写回报告。

#### Scenario: 写回后新增错误值被报告

- **WHEN** 写回后错误值计数大于写回前基线
- **THEN** 报告 SHALL 列出新增错误值的 sheet 清单

### Requirement: 不再执行全表重算

增量合并写回链路 SHALL NOT 触发全表重算（CalculateFull / ensure_formula_cache 全量重算路径从写回链移除）。

#### Scenario: 写回不触发全表重算

- **WHEN** 执行增量合并写回
- **THEN** 链路 SHALL 不调用全表重算，仅依赖 COM Save 的局部重算

### Requirement: 事务备份/回滚链不变

写回失败 SHALL 从备份恢复目标表，报告 SHALL 携带 backup_path 与恢复状态。

#### Scenario: 写回失败回滚

- **WHEN** 写回过程中抛错
- **THEN** 目标表 SHALL 从备份恢复，并抛出带 backup_path / restored 信息的 AddMergeError

### Requirement: 公式写入拒绝

staging 写回计划含公式单元格（`formula` 字段）时 SHALL 报错拒绝，提示改为静态值。

#### Scenario: staging 含公式单元格

- **WHEN** staging 计划中存在公式单元格
- **THEN** 写回 SHALL 抛错拒绝，目标文件不被修改
