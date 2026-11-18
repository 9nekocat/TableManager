## 1. Base

- [x] 1.1 在 `manager_core.py` 定义 `TEMP_DIR = RAW_DIR / ".temp"`

## 2. Diff

- [x] 2.1 修改 `diff_engine.py:export_diff_to_excel` 输出路径使用 `TEMP_DIR`

## 3. Merge

- [x] 3.1 修改 `table_server.py:api_merge_preview` 输出到 `TEMP_DIR`
- [x] 3.2 修改 `table_server.py:api_merge_resolve` 从 `TEMP_DIR` 读取 preview，修正原文件推断逻辑
- [x] 3.3 修改 `table_server.py:api_merge_abort` 从 `TEMP_DIR` 读取并删除 preview
- [x] 3.4 确保 `.temp/` 目录在写入前自动创建（`os.makedirs` 或 `mkdir(parents=True)`）

## 4. Verify

- [x] 4.1 跑现有测试确保无回归
- [x] 4.2 重新打包
