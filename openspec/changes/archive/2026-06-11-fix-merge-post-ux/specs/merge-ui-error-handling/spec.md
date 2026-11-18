# Merge UI Error Handling

Web UI 中合并相关 API 调用的容错处理，确保 spinner 在所有执行路径消失。

## ADDED Requirements

### Requirement: All merge fetch calls have error handlers

前端所有合并相关 fetch 调用 SHALL 具有 `.catch()` 错误处理，失败时恢复 UI 状态并显示错误信息。

#### Scenario: Spinner disappears on network failure
- **WHEN** `loadMergeStatus()` fetch 因网络不可达而失败
- **THEN** UI 显示错误信息（"查询失败: <message>"），spinner 消失

#### Scenario: Spinner disappears on timeout
- **WHEN** `previewMerge()` 或 `resolveMerge()` 的 fetch 超过 30 秒无响应
- **THEN** 请求被 AbortController 取消，UI 显示"操作超时"，spinner 消失

#### Scenario: Spinner disappears on server error  
- **WHEN** API 返回 500 错误
- **THEN** UI 解析 error JSON 并显示错误信息，spinner 消失

### Requirement: loadMergeStatus called only when needed

`loadMergeStatus()` 页面加载时的自动调用 SHALL 在失败时不影响主界面。

#### Scenario: loadMergeStatus fails silently on initial load
- **WHEN** 页面加载 500ms 后自动调用 `loadMergeStatus()` 且失败
- **THEN** merge-panel 区域显示简洁错误但不弹出 alert，不影响表格浏览功能

### Requirement: runSvnUpdate chain handles all states

`runSvnUpdate()` 在 SVN 更新后的链式操作 SHALL 正确处理冲突/非冲突状态。

#### Scenario: Update with conflicts triggers status reload
- **WHEN** SVN 更新返回 `conflicted > 0`
- **THEN** 800ms 后自动调用 `loadMergeStatus()` 加载冲突详情

#### Scenario: Update without conflicts shows clean state
- **WHEN** SVN 更新返回 `conflicted == 0`
- **THEN** 直接显示更新摘要，不再触发 loadMergeStatus

#### Scenario: Update failure shows error
- **WHEN** SVN 更新 API 调用失败
- **THEN** 显示错误信息，恢复按钮状态，不显示 spinner
