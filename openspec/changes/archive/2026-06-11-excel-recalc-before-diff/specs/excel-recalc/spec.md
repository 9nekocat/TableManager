## ADDED Requirements

### Requirement: Excel COM recalculation

`excel_recalc.recalc_xlsx(filepath)` SHALL 用 Excel COM 打开 xlsx 文件，执行全量重算，保存并关闭。

#### Scenario: Recalculate formula cells

- **WHEN** 对包含 XLOOKUP 公式且缓存值为 `#N/A` 的 xlsx 调用 `recalc_xlsx`
- **THEN** 缓存值 SHALL 更新为公式的实际计算结果

#### Scenario: Recalculate with external links

- **WHEN** xlsx 包含跨文件引用（如 `[4]#Main!...`）
- **THEN** Excel SHALL 尝试解析引用并返回正确值（若外部文件不可用则保留 `#REF!`）

### Requirement: Diff recalcs source before comparison

diff 入口函数 SHALL 在比较前对源文件执行重算。

#### Scenario: export_diff_to_excel recalcs current file

- **WHEN** 调用 `export_diff_to_excel`
- **THEN** 当前文件 SHALL 被 copy 到 TEMP_DIR 并重算，diff 基于重算后的副本

#### Scenario: diff_two_files recalcs new file

- **WHEN** 调用 `diff_two_files(old_path, new_path, output_dir)`
- **THEN** `new_path` SHALL 被重算后再参与比较；`old_path` 不重算（保持 SVN 原始状态）
