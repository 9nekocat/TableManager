## Why

当前 TableManager 是 Flask 命令行工具,启动后弹出浏览器窗口,体验不像桌面应用。且缺少拖拽归类、SVN 提交、远端更新提醒等日常高频操作入口,用户需在命令行和浏览器之间反复切换。

## What Changes

- **pywebview 原生窗口**: 用 Edge WebView2 嵌入前端,去掉浏览器跳转和命令行黑窗,体验等同桌面应用
- **拖拽移动分类**: 表格行支持拖拽到侧边栏分类,自动更新 `categories.json` 并刷新视图
- **SVN 一键提交**: 底部面板展示已修改文件列表,勾选 + 输入 commit message 后一键 `svn commit`
- **SVN 远端更新提醒**: 定时查询远端 SVN revision,有更新时 toolbar 角标提示
- **行内智能比对按钮**: 每个表格行新增 [比对] 按钮; 文件已修改→diff vs HEAD; 未修改→diff vs PREV; 复用现有版本比对面板展示结果
- **前端微优化**: 保留原生 JS/CSS, 增强拖拽视觉反馈、过渡动效、空状态、通知条等交互细节

## Capabilities

### New Capabilities

- `native-window`: pywebview 集成,原生桌面窗口替代浏览器,无 console 窗口
- `drag-categorize`: HTML5 拖拽重新归类表格,后端 `categories.json` 原子写入
- `svn-commit`: 图形化 SVN 提交面板,支持文件勾选和自定义 commit message
- `svn-update-notify`: 定时检测 SVN 远端新版本,角标提醒
- `inline-smart-diff`: 表格行内一键智能比对,自动选择 HEAD 或 PREV

### Modified Capabilities

(无——现有 spec 不变)

## Impact

| 层 | 影响 |
|---|------|
| `table_server.py` | 新增 5 个 API 路由; main() 使用 pywebview 启动而非 app.run() |
| `templates/index.html` | 拖拽交互、提交面板、更新角标、行内比对按钮、微动效 |
| `svn_client.py` | 新增 `commit()`, `check_updates()`, `get_file_prev_rev()` |
| `manager_core.py` | 新增 `move_file_to_category()` |
| `diff_engine.py` | 新增 `smart_diff()` 便捷函数 |
| `TableManager.spec` | console=False, hiddenimports 加 pywebview 相关模块 |
| 依赖 | 新增 `pywebview`, `pythonnet` (pywebview Windows 依赖) |
