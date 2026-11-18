## Why

合并功能上线后存在 4 个阻塞性 bug：合并预览中 add/delete 行数据不显示 + 默认决策不正确、合并后无 SVN 权限管理、Web UI 在特定场景下 spinner 不消失、Web API 合并成功后不清理 r*** 临时文件。这些问题导致合并功能基本不可用——用户无法正确预览冲突、合并后文件残留、UI 卡死。

## What Changes

### Bug 1: 合并预览 add/delete 显示 & 默认行为

- `merge_engine.py` `MergePreview.generate()`: 修复 `_write_labeled_row` 在 add/delete 场景下数据显示逻辑。对于 added 行，mine/theirs 侧数据用对应 side 而非 base；对于 deleted 行，确保 base 数据正确回退显示。
- `_write_final_preview_row`: 非冲突行（单方修改/新增/删除）的 `[最终]` 行默认采纳有修改方的数据。当前只对 fill_side 做了正确选择，但具体单元格值在冲突时未按行级状态正确预填。

### Bug 2: 合并后 SVN 权限解锁 & 失败通知

- `table_server.py` `api_merge_resolve`: 新增 SVN lock/unlock 逻辑。合并前 lock 目标文件，成功后 unlock，失败时保留 lock 并返回明确错误信息。
- `table_manager.py` `cmd_merge_resolve`: 同步添加 lock/unlock 流程。
- `merge_engine.py` `MergeResolver.resolve()`: 添加返回值中的错误详情（不再只抛异常）。

### Bug 3: Web UI spinner 不消失

- `templates/index.html`: 修复 `loadMergeStatus()` 默认自动加载时的超时/错误处理；确保所有 fetch 调用有 `.catch()` 使 spinner 在失败时消失；修复 `runSvnUpdate` 冲突后自动加载 merge status 的链式调用。

### Bug 4: Web API 合并后不删除 r*** 文件

- `table_server.py` `api_merge_resolve`: 合并成功后添加清理逻辑，删除 `RAW_DIR` 中的 `*.r*`、`*.mine` 文件（参考 `cmd_merge_resolve` 已有的清理代码）。

## Capabilities

### New Capabilities

- `merge-preview-display`: 合并预览中各行类型的数据正确显示，非冲突行默认采纳修改方
- `merge-post-cleanup`: 合并成功后清理 SVN 冲突产物文件（Web API 端）
- `merge-permission-lock`: SVN 文件锁管理（合并前 lock / 成功后 unlock / 失败保留）
- `merge-ui-error-handling`: Web UI 合并相关 API 调用的容错处理，确保 spinner 在所有路径消失

### Modified Capabilities

无（现有 specs 目录为空）

## Impact

- `merge_engine.py`: `MergePreview.generate()`, `_write_labeled_row`, `_write_final_preview_row`, `MergeResolver.resolve()`
- `table_server.py`: `api_merge_resolve`
- `table_manager.py`: `cmd_merge_resolve`
- `templates/index.html`: `loadMergeStatus`, `runSvnUpdate`, `previewMerge`, `resolveMerge`, `loadTables`
- `svn_client.py`: 可能需要新增 `lock()` / `unlock()` 方法
