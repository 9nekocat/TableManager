## ADDED Requirements

### Requirement: Drag table row to move category
用户 SHALL 能通过拖拽将表格行从一个分类移动到另一个分类。拖拽释放后 `categories.json` 必须原子更新且 UI 立即反映新状态。

#### Scenario: 成功拖拽到侧边栏分类
- **WHEN** 用户拖拽表格行到侧边栏目标分类上并释放
- **THEN** 系统发送 `POST /api/categories/move` 请求,更新 `categories.json`,刷新表格视图,该行出现在目标分类下

#### Scenario: 成功拖拽到主内容区分类 header
- **WHEN** 用户拖拽表格行到主内容区目标分类 header 上并释放
- **THEN** 系统执行与拖拽到侧边栏相同的移动逻辑

#### Scenario: 拖拽到未分类
- **WHEN** 用户拖拽表格行到 "未分类" drop zone
- **THEN** 系统将该文件从原分类移除,添加至 `uncategorized.files`,更新 `categories.json`

#### Scenario: 拖拽到同一分类
- **WHEN** 用户拖拽表格行到其当前所在分类
- **THEN** 系统不发送 API 请求,UI 不变化

#### Scenario: API 失败时回滚
- **WHEN** `POST /api/categories/move` 返回错误
- **THEN** 表格行回到原位置,显示 toast 错误提示 "移动失败: <原因>"

### Requirement: Drag visual feedback
拖拽过程 MUST 提供清晰的视觉反馈。

#### Scenario: 拖拽开始
- **WHEN** 用户开始拖拽表格行
- **THEN** 该行变为半透明 (opacity: 0.4),鼠标旁出现文件名的拖拽幽灵图

#### Scenario: 进入有效 drop zone
- **WHEN** 拖拽中的表格行悬停在分类区域上方
- **THEN** 该分类显示高亮边框 (border-color: accent, border-width: 2px) 和浅色背景

#### Scenario: 离开 drop zone
- **WHEN** 拖拽中的表格行离开分类区域
- **THEN** 高亮边框和背景恢复原样

### Requirement: Atomic category config update
`categories.json` 写入 MUST 使用原子操作,防止配置文件损坏。

#### Scenario: 写入过程中进程崩溃
- **WHEN** 系统在 write 操作中被强制终止
- **THEN** `categories.json` 保持原内容不变 (临时文件未替换)

#### Scenario: 并发写入
- **WHEN** 多个请求同时修改分类
- **THEN** 每个请求使用 `os.replace` 原子替换,最后一个写入者胜出,文件始终完整
