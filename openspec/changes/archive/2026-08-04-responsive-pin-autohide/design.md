## Context

当前 `templates/index.html` 1600+ 行单体文件。新增 3 个窗口交互特性(响应式/置顶/autohide) 加上之前已有的拖拽/提交/比对/面板,状态管理复杂度已饱和。

React 重构同时可完成 UI 美化——利用 `art/charater.png` (角色立绘) 和 `art/Title_Font.ttf` (游戏标题字体) 打造沉浸式 sci-fi 控制台风格。

## Goals / Non-Goals

**Goals:**
- Vite + React 组件化,拆分 8-10 个组件
- React Context 管理全局状态 (categories, SVN, UI mode)
- CSS Modules 保持现有暗色主题设计系统 (CSS variables)
- Canvas 粒子背景 (低开销,不影响表格可读性)
- 角色立绘作为装饰元素 (半透明背景 / 侧边栏 banner)
- 标题字体用于应用头部
- 竖屏/置顶/贴边 autohide 全部整合进新架构
- 构建产物为单目录 (`gui/dist/`), PyInstaller 直接打包

**Non-Goals:**
- 不引入 React Router (单页应用无需)
- 不引入状态管理库 (Context 足够)
- 不改变后端 API

## Decisions

### D1: React 架构 — Vite + React + CSS Modules

```
gui/
├── public/
│   └── art/              ← 构建时从 ../art 复制
│       ├── charater.png
│       └── Title_Font.ttf
├── src/
│   ├── main.jsx
│   ├── App.jsx           ← 根组件, Context Provider
│   ├── App.module.css
│   ├── context/
│   │   └── AppContext.jsx  ← 全局状态 (tables, svnStatus, uiMode)
│   ├── components/
│   │   ├── Sidebar.jsx / .module.css
│   │   ├── Toolbar.jsx / .module.css
│   │   ├── TableGrid.jsx / .module.css
│   │   ├── TableRow.jsx / .module.css
│   │   ├── CommitPanel.jsx / .module.css
│   │   ├── DiffPanel.jsx / .module.css
│   │   ├── ValidationPanel.jsx / .module.css
│   │   ├── MergePanel.jsx / .module.css
│   │   ├── Toast.jsx / .module.css
│   │   └── ParticleBg.jsx / .module.css
│   ├── hooks/
│   │   ├── useApi.js         ← fetch 封装
│   │   ├── useWindowSize.js  ← 响应式检测
│   │   └── useSvnPoll.js     ← SVN 轮询
│   └── styles/
│       └── global.css        ← CSS variables, font-face, reset
├── index.html
├── vite.config.js
└── package.json
```

### D2: 状态设计 — React Context

```js
AppContext = {
  // Data
  categories: [],        // 分类+表格
  svnStatus: null,       // SVN 状态
  mergeStatus: null,     // 合并状态
  updateInfo: null,      // 远端更新

  // UI mode
  narrowMode: false,     // <700px
  pinned: false,         // 大头钉
  autohide: false,       // 贴边隐藏
  activePanel: null,     // 'commit' | 'diff' | 'validation' | 'merge'

  // Actions
  refreshTables: fn,
  togglePin: fn,
  toggleAutohide: fn,
}
```

### D3: 粒子背景 — Canvas 2D

60fps Canvas 全屏背景层 (`position: fixed; z-index: 0; pointer-events: none`):

- ~80 个半透明粒子, 1-2px 直径
- 缓慢上浮 + 随机水平漂移
- 偶发连线 (距离 < 120px 的粒子间画半透明线, 形成网状)
- 配色: `rgba(0, 120, 212, 0.15)` (accent blue) + `rgba(255,255,255,0.04)`
- 使用 `requestAnimationFrame`, 不用 setInterval
- 性能: CPU <2%, GPU 渲染, 不影响表格滚动

CSS `mix-blend-mode: screen` 叠加在暗色背景上, 不遮挡文字。

### D4: UI 视觉美化

**色彩升级** (CSS variables 覆盖):
```
--bg-root: #0a0e14;        (更深,接近纯黑, sci-fi 感)
--bg-sidebar: #0d1117;     (微偏蓝)
--accent: #4fc3f7;         (冰蓝, 原 #0078d4)
--accent-dim: #0d3b4f;
--text: #b0bec5;
--text-bright: #e0eef5;
```

**角色立绘**:
- 侧边栏底部装饰: `position: sticky; bottom: 0;` 半透明角色图 (`opacity: 0.25; mix-blend-mode: soft-light`)
- 不干扰导航文本

**标题字体**:
- `@font-face` 加载 `Title_Font.ttf`
- 侧边栏 header 标题使用: 字号 18px, 字距 2px, 渐变文字 (accent → white)

**微交互**:
- 行 hover: `transform: translateX(2px)` 微位移
- 面板展开: CSS `@keyframes slideUp` 200ms
- 按钮: `scale(0.97)` active 态

### D5: 竖屏 — CSS media query + useWindowSize hook

```css
@media (max-width: 700px) {
  .sidebar { display: none; }
  .main { width: 100%; }
  .table-row { grid-template-columns: 2fr 80px 1fr; } /* 精简列 */
}
```

`useWindowSize` hook 监听 resize → `AppContext.narrowMode` → 组件按 mode 切换渲染。

### D6: 大头钉 — ctypes SetWindowPos

```
POST /api/window/always-on-top  {enabled: bool}
→ ctypes.windll.user32.SetWindowPos(hwnd, HWND_TOPMOST, ...)
```

pywebview 6.x 通过 `webview.windows[0]._hwnd` 获取原生句柄 (Windows 特化)。

### D7: 贴边 autohide — daemon 线程 + GetWindowRect

200ms 间隔轮询线程 (仅 autohide 启用时), 检测贴边+鼠标位置, `SetWindowPos` 滑入/滑出。

## Risks

| Risk | Mitigation |
|------|-----------|
| Vite build 产物与 PyInstaller 路径不一致 | `vite build --outDir dist --base ./` 相对路径 |
| React 打包后首次加载慢 | 代码 split 不需要 — 整个 app 很小; gzip < 50KB JS |
| 粒子 Canvas 在集成显卡上掉帧 | 降低粒子数到 40, 关闭连线, `will-change: transform` |
| 角色图 741KB 拖慢加载 | 构建时压缩 (sharp/vite-imagemin), 目标 < 150KB WebP |
