## Why

合并决议中删除行后，输出文件留空行（gap）。根因是 `row_map` 用 `old_row + insert_offset` 计算，只处理插入引起的下移，不处理删除引起的上移。被删行在 Phase 4 被清为 None，但其输出位置仍被 row_map 保留，后续行不往上挤。

## What Changes

- Phase 3 构建 `ordered` 时跳过被删行
- 用顺序编号替代 offset 公式生成 `row_map`
- 去掉 Phase 4 的清空所有 cell 逻辑和逐行清除逻辑
- 尾部多余行清理

## Impact

- `merge_engine.py`：Phase 2/3/4（`resolve` 方法内）
- 不碰检测（ConflictDetector）、预览（MergePreview）
