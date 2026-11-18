## 1. 预览生成 — 新增修改类型列

- [x] 1.1 `_write_labeled_row` 新增 `change_type_label` 参数，在 B 列写入变更类型标注
- [x] 1.2 `_write_final_preview_row` 新增 `default_decision` 参数，在 B 列写入下拉默认值
- [x] 1.3 `generate` 方法：调整所有写行调用，传入正确的 change_type_label 和 default_decision
- [x] 1.4 `generate` 方法：所有数据列偏移 +1（`excel_col = col_idx + 2`）
- [x] 1.5 `generate` 方法：在 [最终] 行 B 列区域添加 DataValidation（下拉：`[修改],[删除]`）
- [x] 1.6 MINE_DELETED 改为写 `[最终]` 行（之前跳过），默认值 `[删除]`
- [x] 1.7 THEIRS_DELETED [最终] 行默认值从 mine 数据改为 `[删除]`

## 2. 决议器 — 读取修改类型列

- [x] 2.1 `resolve` 方法：解析时读取 [最终] 行 B 列值作为决策
- [x] 2.2 去掉 `group_has_mine` + `base_order` 隐式 3-way 分类逻辑
- [x] 2.3 THEIRS_ADDED / BOTH_ADDED 新增行的插入位置算法保持锚点搜索，但只对 B 列为 `[修改]` 的行计算插入
- [x] 2.4 全部数据列读取偏移 +1（读取 C 列起，跳过 A 列标记和 B 列修改类型）

## 3. 测试验证

- [x] 3.1 手动生成合并预览，验证所有冲突类型的修改类型列标注正确
- [x] 3.2 测试下拉功能在 Excel 中正常工作
- [x] 3.3 测试决议器：[修改] 选项保留行、[删除] 选项删除行
- [x] 3.4 测试 THEIRS_DELETED 默认删除行为
- [x] 3.5 测试 MINE_DELETED 显式删除并可改为保留
- [x] 3.6 回归测试：单方修改（MINE_MOD / THEIRS_MOD）正确应用
