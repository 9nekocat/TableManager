## Context

TableManager 的合并功能通过 CLI (`table_manager.py`) 和 Web UI (`table_server.py` + `templates/index.html`) 两个入口暴露。核心引擎在 `merge_engine.py`（冲突检测、预览生成、合并解析）。

当前状态：
- **合并预览**: `MergePreview.generate()` 对不同的 row status（MINE_ADDED, THEIRS_DELETED 等）生成 `[我的]`/`[仓库]`/`[最终]` 三行。其中 add/delete 行的数据源选择逻辑有边界 bug。
- **SVN 文件锁**: `SvnClient` 没有 `lock()`/`unlock()` 方法。合并操作直接写入文件，不管理权限。
- **Web API 清理**: `api_merge_resolve` 合并成功后调用 `svn resolved` 但不删除 r***/.mine 文件（CLI 版本 `cmd_merge_resolve` 有清理逻辑）。
- **UI 容错**: 前端 `index.html` 的 fetch 调用缺少完整的错误处理链，spinner 在 API 不可达/超时时不消失。

## Goals / Non-Goals

**Goals:**
- Bug 1: 修复 `_write_labeled_row` 在 add/delete 行的数据显示；`_write_final_preview_row` 为非冲突行默认采纳修改方的数据
- Bug 2: `SvnClient` 新增 `lock()`/`unlock()`；合并流程集成 lock→resolve→unlock
- Bug 3: 前端所有 fetch 调用添加 `.catch()` 和超时处理，确保 spinner 在所有路径消失
- Bug 4: `api_merge_resolve` 添加 r***/.mine 文件清理

**Non-Goals:**
- 不改动 ConflictDetector 的冲突检测逻辑
- 不新增批量合并功能
- 不修改 diff_engine.py

## Decisions

### Decision 1: add/delete 行数据显示策略

**选择**: 在 `_write_labeled_row` 和 `_write_final_preview_row` 中分别按 row status 决定数据源。

- `MINE_ADDED`: mine 侧显示 mine 数据 → `[最终]` 默认 mine
- `THEIRS_ADDED`: theirs 侧显示 theirs 数据 → `[最终]` 默认 theirs
- `MINE_DELETED`: mine 侧显示 base 数据（"你删了什么"）→ `[最终]` 默认 `[删除]`
- `THEIRS_DELETED`: theirs 侧显示 base 数据 → `[最终]` 默认 `[删除]`
- `CONFLICT` / `BOTH_ADDED`: 双方数据都显示 → `[最终]` 需要用户选择

**替代方案**: 对所有行类型统一用 "last writer wins" 自动合并 → 不可取，用户需要预览冲突。

### Decision 2: SVN lock/unlock 集成位置

**选择**: 在 `table_server.py` `api_merge_resolve` 和 `table_manager.py` `cmd_merge_resolve` 中分别添加 lock/unlock 逻辑（调用 `SvnClient`）。

lock 在 resolve 前执行，resolve 成功后 unlock，resolve 失败保留 lock 并通知用户手动处理。

**替代方案**: 在 `MergeResolver.resolve()` 内部管理 → 不可取，lock 是纯 SVN 操作，不应耦合到合并引擎。

### Decision 3: 前端 spinner 容错

**选择**: 给所有 fetch 调用添加：
1. `.catch()` 块，显示错误信息并恢复 UI 状态
2. `AbortController` + 30 秒超时
3. `loadMergeStatus()` 默认自动调用改为仅在页面加载成功后可选触发（移除空跑 spinner）

**替代方案**: 全局 fetch 拦截器 → 过度设计，4 个调用点不值得。

### Decision 4: r*** 文件清理复用

**选择**: 将 `cmd_merge_resolve` 中的清理逻辑提取为独立函数 `cleanup_conflict_artifacts()`，在 CLI 和 Web API 两端复用。

## Risks / Trade-offs

- [SVN lock 在网络异常时可能未释放] → 在失败路径中明确告知用户文件路径，并提示 `svn unlock` 手动命令
- [r*** 文件清理可能误删] → glob 模式精确匹配 `{stem}.*.r*`，只删除当前文件的冲突产物
- [30 秒超时对大型合并预览可能不够] → MergePreview 生成耗时与表格大小相关，超时值可配置，默认 30s 足够覆盖当前表格规模
