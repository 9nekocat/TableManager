## 1. 核心优化

- [x] 1.1 `_get_old_version()` 去掉 Excel recalc，直接读 SVN temp 文件
- [x] 1.2 `generate_diff()` 去掉 `recalc_to_temp`，直接读工作副本
- [x] 1.3 `export_diff_to_excel()` 去掉 `recalc_to_temp`，直接读工作副本

## 2. 批量重算

- [x] 2.1 `excel_recalc.py` 新增 `recalc_batch()` 函数
- [x] 2.2 `diff_two_files()` 改用 copy-to-temp + `recalc_batch`

## 3. 清理

- [x] 3.1 移除 `diff_engine.py` 中未使用的 `recalc_to_temp` 导入
- [x] 3.2 恢复 `_cell_display()` 为纯值显示

## 4. 验证

- [x] 4.1 Effect.xlsx diff vs HEAD: 10.0s → 1.1s
- [x] 4.2 diff_two_files(Effect_prev, Effect_latest): 带批量重算 5.9s，输出含实际值
- [x] 4.3 回归测试 14/14 PASS
