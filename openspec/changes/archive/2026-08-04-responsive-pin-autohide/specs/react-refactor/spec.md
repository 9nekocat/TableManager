## ADDED Requirements

### Requirement: Vite + React component architecture
前端 MUST 使用 Vite + React 重建,拆分为独立组件和 CSS Modules,不再使用单体 HTML 文件。

#### Scenario: 构建产物
- **WHEN** 执行 `npm run build`
- **THEN** `gui/dist/` 目录生成,包含 `index.html` + `assets/`,可被 PyInstaller 直接打包为 Flask 模板

#### Scenario: 开发模式
- **WHEN** 执行 `npm run dev`
- **THEN** Vite dev server 启动于 localhost:5173, `/api` 请求 proxy 到 Flask 后端

#### Scenario: 组件独立修改
- **WHEN** 修改 `Sidebar.jsx` 的渲染逻辑
- **THEN** 其他组件 (Toolbar, TableGrid) 不受影响,无需重新测试

### Requirement: React Context state management
全局状态 MUST 通过 React Context 管理,包括: 分类表格数据、SVN 状态、合并状态、UI 模式 (窄屏/大头钉/autohide/活跃面板)。

#### Scenario: 拖拽移动分类刷新
- **WHEN** 拖拽 API 返回成功
- **THEN** AppContext.refreshTables() 触发,所有依赖组件自动重渲染

#### Scenario: 面板切换
- **WHEN** 用户点击 toolbar 按钮
- **THEN** AppContext.activePanel 更新,对应面板组件展示,旧面板卸载

### Requirement: Backward-compatible API consumption
所有现有 Flask API 路由 MUST 保持不变,React 前端使用相同的 API 端点和数据格式。

#### Scenario: loadTables API
- **WHEN** App 组件挂载
- **THEN** 调用 `GET /api/tables`,解析返回的 category 分组 JSON,填充 Context

### Requirement: All existing features preserved
React 重构 MUST 完整保留以下功能: 表格浏览/搜索/打开 Excel、校验、SVN 状态/更新、合并预览/解决、版本比对(手动+智能)、拖拽分类、提交面板、更新提醒、split-button 比对。

#### Scenario: 功能完整性
- **WHEN** 用户使用 React 版本
- **THEN** 每个现有功能的操作流程与原生版本一致,响应结果相同
