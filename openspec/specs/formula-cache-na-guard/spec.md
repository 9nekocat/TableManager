# formula-cache-na-guard Specification

## Purpose
定义写回前后 #N/A / #REF! 错误值基线校验的行为契约：写回前统计各 # sheet 的错误值缓存计数作为基线，写回后重新统计并对比，新增错误值即报告，并区分链接问题与表格数据问题。
## Requirements
### Requirement: 写回前记录错误值基线

写回执行前 SHALL 统计目标表所有 # sheet 的错误值（#N/A、#REF! 等）缓存计数，作为对比基线。

#### Scenario: 基线统计覆盖全部数据 sheet

- **WHEN** 目标表包含多个 # 开头的 sheet
- **THEN** 基线 SHALL 按 sheet 分别统计错误值计数并给出总数

### Requirement: 写回后对比基线并报告新增错误值

写回成功后 SHALL 重新统计错误值并与基线对比：新增错误值 SHALL 报告（列出涉及 sheet 供策划处理）；基线已存在的错误值 SHALL 视为表格数据问题单独列出，不判定为写回失败。

#### Scenario: 无新增错误值

- **WHEN** 写回后错误值计数与基线一致
- **THEN** 校验 SHALL 通过，报告不包含新增错误值

#### Scenario: 出现新增错误值

- **WHEN** 写回后某 sheet 错误值计数大于基线
- **THEN** 校验 SHALL 失败，并报告新增错误值数量与涉及 sheet 清单

### Requirement: 校验结果进入写回报告

写回报告的返回值 SHALL 包含错误值基线校验结果（基线计数、写回后计数、新增数量与涉及 sheet），供 CLI/GUI 展示。

#### Scenario: 报告包含校验数据

- **WHEN** 写回完成且返回报告
- **THEN** 报告 SHALL 包含 baseline / after / added_sheets / ok 字段

