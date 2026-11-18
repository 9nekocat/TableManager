## ADDED Requirements

### Requirement: Periodic SVN remote revision check
系统 SHALL 定期检测 SVN 远端是否有新版本,并在 UI 上显示角标提醒。

#### Scenario: 检测到远端新版本
- **WHEN** 定时检测发现 `remote_rev > local_rev`
- **THEN** toolbar SVN 按钮显示角标数字 (红色圆点或数字表示落后版本数),hover 提示 "远端有新版本 rXXXX,落后 N 个版本"

#### Scenario: 版本已同步
- **WHEN** `remote_rev == local_rev`
- **THEN** SVN 按钮角标消失,回归正常样式

#### Scenario: 网络不可用时静默降级
- **WHEN** `svn info -r HEAD` 超时或网络错误
- **THEN** 系统记录日志,UI 不显示任何错误,保留上次角标状态不变

#### Scenario: 非 SVN 工作副本
- **WHEN** DataTable/Raw 目录不是 SVN 工作副本
- **THEN** 系统跳过所有更新检测,不显示角标

### Requirement: Check scheduling
更新检测 MUST 在启动后首次执行,之后按固定间隔轮询。

#### Scenario: 启动后延迟首次检测
- **WHEN** 应用启动完成
- **THEN** 2 秒后执行首次 `GET /api/svn/check-updates`,之后每 5 分钟执行一次

#### Scenario: 用户手动刷新
- **WHEN** 用户点击 SVN 按钮或刷新按钮
- **THEN** 立即执行一次检测并重置定时器间隔

### Requirement: API timeout protection
远端版本查询 MUST 有超时保护,防止阻塞线程。

#### Scenario: SVN 命令超时
- **WHEN** `svn info -r HEAD` 执行超过 15 秒
- **THEN** 请求被终止,返回 `{timeout: true, has_update: false}`, 前端静默处理
