## Why

当前 `index.html` 超 1600 行, CSS/HTML/JS 全部混在一个文件,全局状态散落,改一处功能需在三处修改。交互复杂度(拖拽/下拉/面板切换/响应式/贴边隐藏)已超出原生 JS 舒适区。同时 UI 纯工具风格,缺乏游戏项目应有的视觉辨识度。

## What Changes

1. **React 重构**: Vite + React + CSS Modules 组件化, 从单文件拆为 8-10 个组件, 状态管理用 React Context
2. **竖屏适配**: 窗口宽度 < 700px 时自动隐藏侧边栏,扁平列表展示全表格, min_size 降至 400
3. **大头钉模式**: toolbar 图钉按钮,切换窗口 `alwaysOnTop`
4. **贴边 autohide**: 窗口贴靠屏幕边时鼠标移出隐藏/移入展开 (4px 抓手)
5. **UI 美化**: 集成游戏角色立绘 (`art/charater.png`) 和标题字体 (`art/Title_Font.ttf`); Canvas 粒子背景(低性能开销,保证可读性); dark sci-fi 调色板升级

## Capabilities

### New Capabilities

- `react-refactor`: Vite + React 组件化重构, Context 状态管理
- `responsive-sidebar`: 窄屏自动隐藏侧边栏,扁平列表
- `always-on-top`: 大头钉窗口置顶
- `edge-autohide`: 贴边鼠标移出隐藏/移入展开
- `game-ui-theme`: 角色立绘 + 标题字体 + 粒子背景 + sci-fi 视觉

### Modified Capabilities

(无——所有现有功能保留,仅重构)

## Impact

| 层 | 影响 |
|---|------|
| 新建 `gui/` | Vite + React 项目: `src/`, `public/`, `vite.config.ts` |
| `templates/index.html` | **移除** — 替换为 `gui/dist/` 构建产物 |
| `table_server.py` | 新增 window API (`always-on-top`, `autohide`); 静态文件路径改为 `gui/dist/` |
| `TableManager.spec` | datas: `('gui/dist', 'templates')`; 新增 `('art', 'art')` 打包素材 |
| `art/` | 角色立绘 + 标题字体 — 打包进 EXE |
