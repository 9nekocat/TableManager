## Why

`_diff` 和 `_merge` 临时文件散落在 `RAW_DIR` 下，与源数据文件混在一起。集中到 `.temp/` 便于 `.svnignore` 和清理。

## What Changes

- `diff_engine.py`：`export_diff_to_excel` 输出到 `RAW_DIR / ".temp" / <relpath>`
- `table_server.py`：`api_merge_preview` 输出到 `.temp/`，`api_merge_resolve` 和 `api_merge_abort` 从 `.temp/` 读写
- 路径用 `RAW_DIR` 相对路径构建，不硬编码绝对路径

## Capabilities

### New Capabilities

- `temp-artifact-path`: 所有临时产物（diff、merge preview）统一输出到 `RAW_DIR/.temp/` 下

### Modified Capabilities

<!-- none -->

## Impact

- `diff_engine.py`：`export_diff_to_excel` 输出路径
- `table_server.py`：`api_merge_preview`、`api_merge_resolve`、`api_merge_abort` 路径逻辑
