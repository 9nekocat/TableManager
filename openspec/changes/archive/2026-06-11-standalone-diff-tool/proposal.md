## Why

TortoiseSVN 右键比对 DataTable xlsx 文件时，内置 diff 无法识别公式语义。需要独立 diff exe，被 TortoiseSVN 调用快速比对。

## What Changes

- `diff_engine.py`：新增 `diff_two_files(old_path, new_path, output_dir)` — 直接比较两个文件，不走 SVN
- `diff_standalone.py`：新增 CLI 入口，接收 `%base %mine` 两个参数
- `diff_standalone.spec`：新增轻量 PyInstaller 配置

## Capabilities

### New Capabilities

- `standalone-diff`: 独立 diff exe，可被 TortoiseSVN 调用

## Impact

- `diff_engine.py`：新增函数
- `diff_standalone.py`：新文件
- `diff_standalone.spec`：新文件
