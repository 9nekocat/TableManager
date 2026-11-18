# Diff Temp Cleanup

diff 临时文件唯一命名 + 异常安全清理。

## ADDED Requirements

### Requirement: Unique temp file names per diff run

`diff_two_files()` SHALL 为每次调用生成唯一的临时文件名，使用 UUID 后缀避免跨次运行冲突。

#### Scenario: Two consecutive diff runs use different temp names
- **WHEN** 连续两次调用 `diff_two_files(old, "Raw/Effect.xlsx", out_dir)`
- **THEN** 第一次的临时文件名为 `Raw/Effect_old_recalc_a1b2c3.xlsx`
- **THEN** 第二次的临时文件名为 `Raw/Effect_old_recalc_d4e5f6.xlsx`
- **THEN** 两次运行的文件路径不同，互不干扰

### Requirement: Exception-safe temp cleanup

`diff_two_files()` SHALL 在异常路径也删除临时文件。

#### Scenario: recalc_batch fails, temps still cleaned
- **WHEN** `recalc_batch` 抛出异常
- **THEN** 已创建的临时文件在 finally 块中被删除

#### Scenario: Successful run cleans temps
- **WHEN** diff 成功完成
- **THEN** 两个临时文件在 finally 块中被删除
