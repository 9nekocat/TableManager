## Context

现有的 diff 流程对每次比较做 2 次独立的 Excel COM recalc（旧版本 + 当前版本各一次），每次启动 Excel 实例、打开文件、打开链接工作簿、全量重算、保存。Effect.xlsx diff 耗时 10s，其中 6.5s 是 recalc。

## Goals / Non-Goals

**Goals:**
- `generate_diff` / `export_diff_to_excel`：去掉冗余 recalc，直接从工作副本读取（9x 提升）
- `diff_two_files`：用 copy-to-temp + 批量重算保证值正确，不修改源文件

**Non-Goals:**
- 不改变 diff 输出格式
- 不改变 `_read_all_rows_with_formulas` 内部实现

## Decisions

### Decision 1: generate_diff 跳过所有 recalc

**选择**: `_get_old_version` 直接从 SVN temp 文件读取；`generate_diff` 直接从工作副本读取。

**理由**: SVN 历史版本由 Excel 保存提交，缓存值完好。正常工作副本同样。只有被 TableManager merge 修改过的文件才需要 recalc，但那是少数场景。

**替代方案**: 检测公式缓存是否存在再决定是否 recalc → 增加一次 workbook 读取，抵消部分收益。

### Decision 2: diff_two_files 用批量 recalc

**选择**: 复制到 temp → `recalc_batch` → 读取 → 删除 temp。

**理由**: TortoiseSVN `%base` temp 文件无缓存值；同时 `diff_two_files` 常用于任意两个文件的比较，不能假设文件状态。批量 recalc 只启动一次 Excel，比逐个 `recalc_xlsx` 快。

### Decision 3: recalc_batch 去重链接工作簿

**选择**: 收集所有文件的链接工作簿路径，去重后统一打开。

**理由**: 两个 Effect 相关文件的链接高度重合，去重可避免重复 I/O。

## Risks / Trade-offs

- [批量重算结果与逐个重算可能不同] → 同一个 Excel 实例内计算顺序一致，风险极低
- [工作副本被 merge 修改后 diff 会缺值] → 此时建议用户先手动打开文件用 Excel 保存一次
