## 1. React 项目搭建

- [x] 1.1 `gui/` 目录: `npm create vite@latest gui -- --template react` 初始化
- [ ] 1.2 `vite.config.js`: 配置 `base: './'`, proxy `/api` → `http://127.0.0.1:5000`, build `outDir: 'dist'`
- [ ] 1.3 `global.css`: CSS variables (sci-fi 调色板), `@font-face Title_Font`, reset
- [ ] 1.4 `App.jsx` + `AppContext.jsx`: 根组件 + Context (categories, svnStatus, mergeStatus, updateInfo, uiMode)
- [ ] 1.5 `hooks/useApi.js`: fetch 封装 (统一 base URL, timeout, JSON parse)
- [ ] 1.6 复制 `art/` 到 `gui/public/art/` (charater.png, Title_Font.ttf)

## 2. 核心组件

- [x] 2.1 `Sidebar.jsx`: 分类导航 + 角色立绘底部装饰 + 标题字体 header + 拖拽 drop zone
- [ ] 2.2 `Toolbar.jsx`: 搜索框 + 图钉按钮 + SVN 按钮 + 提交按钮 + 校验按钮 + 版本比对按钮 + 更新角标
- [ ] 2.3 `TableGrid.jsx`: 表格网格 — 从 Context 取数据,分 category section 渲染
- [ ] 2.4 `TableRow.jsx`: 单行 — draggable + [打开] + [比对] split-button + 拖拽 dragstart
- [ ] 2.5 `CommitPanel.jsx` + `DiffPanel.jsx` + `ValidationPanel.jsx` + `MergePanel.jsx`
- [ ] 2.6 `Toast.jsx`: 底部居中 toast,3 秒自动消失, Context 驱动

## 3. 视觉美化

- [x] 3.1 `ParticleBg.jsx`: Canvas 全屏背景粒子 (80 粒子 + 近距离连线, requestAnimationFrame)
- [ ] 3.2 sci-fi 调色板: CSS variables 覆盖 (背景 `#0a0e14`, 主色 `#4fc3f7`)
- [ ] 3.3 角色立绘: 侧边栏底部 `opacity: 0.25; mix-blend-mode: soft-light`
- [ ] 3.4 标题字体: `@font-face` 加载, 侧边栏 header 渐变文字
- [ ] 3.5 微交互: 行 hover `translateX(2px)`, 面板 slideUp 动画, 按钮 scale(0.97) active

## 4. 竖屏适配

- [x] 4.1 `useWindowSize.js` hook: 监听 resize, 更新 `narrowMode` context
- [x] 4.2 CSS `@media (max-width: 700px)`: 侧边栏隐藏, 表格列简化
- [x] 4.3 `TableGrid.jsx` 按 `narrowMode` 选择扁平列表 vs 分类分组渲染
- [x] 4.4 `min_size` 从 900 降到 400 (python table_server.py)

## 5. 大头钉模式

- [x] 5.1 `POST /api/window/always-on-top` API: ctypes SetWindowPos HWND_TOPMOST 切换
- [x] 5.2 `_get_hwnd()`: 通过 webview 获取原生窗口句柄
- [x] 5.3 Toolbar 图钉按钮: 点击 toggle → fetch API → 更新按钮 active 样式

## 6. 贴边 autohide

- [x] 6.1 `POST /api/window/autohide` API: enabled toggle → 启动/停止 daemon 线程
- [x] 6.2 `_autohide_loop()`: 200ms 间隔, GetWindowRect + GetCursorPos, SetWindowPos 滑入/滑出
- [x] 6.3 贴左边缘: rect.left <= 0, 鼠标 x > rect.right + 20 → 隐藏到 -width + 4
- [x] 6.4 贴右边缘: rect.right >= screenWidth, 鼠标 x < rect.left - 20 → 隐藏到 screenWidth - 4
- [x] 6.5 Toolbar autohide 按钮: 点击 toggle, active/inactive 样式

## 7. 构建与打包

- [x] 7.1 `npm run build` → `gui/dist/` 验证
- [x] 7.2 更新 `table_server.py`: 静态文件路径改为 `gui/dist/`
- [x] 7.3 更新 `TableManager.spec`: datas `('gui/dist', 'templates')`, `('art', 'art')`
- [x] 7.4 PyInstaller 打包并验证 EXE

## 8. 验证

- [x] 8.1 React dev 模式: `npm run dev` + `python table_server.py --web` 全功能测试
- [x] 8.2 窄屏切换: 拖拽 <700px / >700px
- [x] 8.3 图钉 toggle, 贴边隐藏/展开
- [x] 8.4 粒子背景性能 (任务管理器 GPU/CPU)
- [x] 8.5 角色立绘 + 标题字体显示
- [x] 8.6 拖拽分类 / 提交面板 / 智能比对 功能完整性
- [x] 8.7 PyInstaller 打包 → 双击 EXE → 完整验证
