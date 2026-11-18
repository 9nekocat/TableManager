# merge-resolve Specification

## Purpose
定义 merge resolve / recover 写回链路的行为契约：合并结果写回改走 COM 直写（局部重算而非全表重算），写回后执行错误值基线校验，事务备份/回滚链保持不变，外部链接与动态数组元数据无需 zip 注入恢复。
## Requirements
### Requirement: resolve 写回走 COM 直写

merge resolve 写回 SHALL 使用 COM 直写应用 [最终] 决策；全表重算（`_refresh_formula_cache` / `ensure_formula_cache`）SHALL 从 resolve/recover 链路移除，仅依赖 COM Save 的局部重算。

#### Scenario: resolve 写回保留缓存

- **WHEN** 应用 merge 预览的 [最终] 决策
- **THEN** 写回 SHALL 通过 COM 直写完成，外部引用缓存不新增 #N/A，公式保持公式

#### Scenario: 无 Excel 环境降级

- **WHEN** 本机无 win32com/Excel
- **THEN** 写回 SHALL 降级 openpyxl 路径并明确标记风险

### Requirement: 写回后错误值基线校验

resolve / recover 写回成功后 SHALL 执行错误值基线校验，新增错误值报告到结果 dict。

#### Scenario: 新增错误值被报告

- **WHEN** resolve 写回后错误值计数大于写回前基线
- **THEN** 结果 SHALL 列出新增错误值的 sheet 清单

### Requirement: 事务备份/回滚链不变

resolve/recover 写回失败 SHALL 从备份恢复，并抛出带 backup_path / restored 信息的 MergeWorkflowError。

#### Scenario: 写回失败回滚

- **WHEN** resolve 写回过程中抛错
- **THEN** 输出表 SHALL 从备份恢复，并抛出 MergeWorkflowError

