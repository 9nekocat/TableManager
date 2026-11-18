## ADDED Requirements

### Requirement: Game character art integration
角色立绘 `art/charater.png` MUST 作为 UI 装饰元素,不干扰表格可读性。

#### Scenario: 侧边栏角色展示
- **WHEN** 应用启动且侧边栏可见
- **THEN** 侧边栏底部显示半透明角色立绘 (`opacity: 0.25, mix-blend-mode: soft-light`),不遮挡分类导航文字

#### Scenario: 窄屏模式无角色图
- **WHEN** 窗口宽度 < 700px
- **THEN** 角色立绘不显示 (侧边栏已隐藏)

#### Scenario: 角色图加载失败
- **WHEN** 角色图片文件无法加载
- **THEN** 侧边栏正常显示,无装饰图,无报错

### Requirement: Title font for app header
标题字体 MUST 用于侧边栏 header 的应用名称。

#### Scenario: 标题渲染
- **WHEN** 应用启动
- **THEN** 侧边栏顶部 "DATATABLE MANAGER" 使用 `Title_Font.ttf`,字号 18px,字距 2px,渐变文字 (accent→white)

#### Scenario: 字体回退
- **WHEN** 标题字体文件加载失败
- **THEN** 回退到系统等宽字体,保持可读

### Requirement: Canvas particle background
系统 MUST 提供 Canvas 粒子背景动画,不影响内容可读性。

#### Scenario: 粒子渲染
- **WHEN** 应用启动
- **THEN** 全屏 Canvas 显示 ~80 个缓慢上浮的半透明粒子 (1-2px),粒子间近距离 (<120px) 有半透明连线

#### Scenario: 不影响交互
- **WHEN** 粒子 Canvas 渲染中
- **THEN** `pointer-events: none`, `z-index: 0`,表格/按钮/面板正常可点击

#### Scenario: 性能安全
- **WHEN** 粒子动画运行
- **THEN** `requestAnimationFrame` 驱动,CPU 占用 <2%,不阻塞 UI 线程

### Requirement: Sci-fi color palette
整体配色 MUST 升级为 sci-fi 暗色风格,保持文本对比度 WCAG AA (4.5:1)。

#### Scenario: 色彩方案
- **WHEN** 应用启动
- **THEN** 背景 `#0a0e14`, 主色调冰蓝 `#4fc3f7`, 文本 `#b0bec5`, 亮文本 `#e0eef5`

### Requirement: Micro-interactions
UI MUST 有微交互增强操作反馈。

#### Scenario: 行 hover
- **WHEN** 鼠标悬停表格行
- **THEN** 行 `transform: translateX(2px)`, 过渡 150ms

#### Scenario: 面板展开
- **WHEN** 底部面板显示
- **THEN** 面板从底部滑入 (`@keyframes slideUp 200ms`)

#### Scenario: 按钮按下
- **WHEN** 用户点击按钮
- **THEN** 按钮 `transform: scale(0.97)`, 过渡 100ms
