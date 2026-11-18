## 1. Bug 1: 合并预览 add/delete 显示 & 默认行为

- [x] 1.1 创建测试用例：MINE_ADDED 行预览数据显示验证
- [x] 1.2 创建测试用例：THEIRS_DELETED 行预览数据显示验证
- [x] 1.3 修复 `MINE_DELETED` fill_side 从 "mine" → "base"（mine 侧无数据）
- [x] 1.4 修复 `_write_final_preview_row` 非冲突行默认决策逻辑
- [x] 1.5 运行测试验证 Bug 1 修复 (14/14 PASS)

## 2. Bug 2: SVN 权限锁管理

- [x] 2.1 `SvnClient` 添加 `lock()` / `unlock()` 方法
- [x] 2.2 创建 SVN lock/unlock 单元测试
- [x] 2.3 `api_merge_resolve` 集成 lock → merge → unlock 流程
- [x] 2.4 `cmd_merge_resolve` 集成 lock → merge → unlock 流程
- [x] 2.5 运行测试验证 Bug 2 修复

## 3. Bug 3: Web UI spinner 不消失

- [x] 3.1 分析现有 fetch 调用的 catch 覆盖情况
- [x] 3.2 所有 merge 相关 fetch 添加 `fetchWithTimeout` (30s 默认超时)
- [x] 3.3 `loadMergeStatus()` 初始加载失败静默处理
- [x] 3.4 手动测试 UI 各场景（断网、超时、500 错误）

## 4. Bug 4: Web API 合并后不删除 r*** 文件

- [x] 4.1 提取 `cleanup_conflict_artifacts()` 到 `manager_core.py`
- [x] 4.2 共享函数清理 *.r* 和 *.mine 文件
- [x] 4.3 `api_merge_resolve` 调用共享清理函数
- [x] 4.4 `cmd_merge_resolve` 改用共享清理函数
- [x] 4.5 运行测试验证 Bug 4 修复

## 5. 归档已完成的 change

- [x] 5.1 归档 `excel-recalc-before-diff`
- [x] 5.2 归档 `fix-deleted-row-gap`
- [x] 5.3 归档 `fix-diff-formula-false-positive`
- [x] 5.4 归档 `fix-merge-data-leak`
- [x] 5.5 归档 `merge-preview-change-type-column`
- [x] 5.6 归档 `standalone-diff-tool`
- [x] 5.7 归档 `temp-dir-for-artifacts`
