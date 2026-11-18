## ADDED Requirements

### Requirement: Inline smart diff button on table rows
每个表格行 SHALL 显示 [比对] 按钮,点击后自动选择比较目标并显示差异结果。

#### Scenario: 已修改文件 diff vs HEAD
- **WHEN** 用户点击某行的 [比对] 按钮且该文件 SVN 状态为 M (已修改)
- **THEN** 系统执行 diff vs HEAD (仓库最新版本),底部 diff panel 展开,显示变更详情

#### Scenario: 未修改文件 diff vs PREV
- **WHEN** 用户点击某行的 [比对] 按钮且该文件 SVN 状态为未修改
- **THEN** 系统获取 `svn log -l 2`,取倒数第二次修改的版本号,执行 diff vs PREV,底部 diff panel 展开

#### Scenario: 文件无历史版本
- **WHEN** 文件在 SVN 中只有一个版本或无历史
- **THEN** 系统提示 "该文件无历史版本可供比对"

#### Scenario: 文件不在 SVN 版本控制下
- **WHEN** 文件是新文件或不在 SVN 中
- **THEN** 系统提示 "该文件不在 SVN 版本控制中"

### Requirement: Diff result display
智能比对结果 SHALL 复用现有版本比对面板 (`diff-panel`) 展示,不创建新的结果区域。

#### Scenario: 结果展示
- **WHEN** diff 比对完成
- **THEN** diff-panel 底部面板展开,渲染变更行/单元格差异,自动滚动到面板可见

#### Scenario: 比对进行中
- **WHEN** diff 请求正在处理
- **THEN** 对应的 [比对] 按钮显示加载状态 (spinner 或 disabled),面板显示 "正在比对..."

### Requirement: Smart diff API
`GET /api/diff/smart/<path>` SHALL 自动判断比对目标版本。

#### Scenario: 自动选择 HEAD
- **WHEN** API 被调用且 `svn status <file>` 返回 M
- **THEN** 执行 `svn diff --diff-cmd ...` vs HEAD, 返回 diff 结果和 `compared_against: "HEAD"`

#### Scenario: 自动选择 PREV
- **WHEN** API 被调用且文件未修改
- **THEN** 执行 diff vs 倒数第二次修改的版本号, 返回 diff 结果和 `compared_against: <revision>`

#### Scenario: API 超时
- **WHEN** SVN diff 操作超过 30 秒
- **THEN** 返回 `{error: "比对超时,请稍后重试"}`, HTTP 500
