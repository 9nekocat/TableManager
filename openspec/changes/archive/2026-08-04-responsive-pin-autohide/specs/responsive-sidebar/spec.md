## ADDED Requirements

### Requirement: Sidebar auto-hide on narrow window
窗口宽度 < 700px 时,系统 SHALL 自动隐藏左侧边栏,表格以扁平列表展示(无分类 section 分组)。

#### Scenario: 窗口缩窄触发隐藏
- **WHEN** 用户拖拽窗口边缘使宽度 < 700px
- **THEN** 侧边栏隐藏 (`display: none`),主内容区占满全宽,表格按文件名排序的扁平列表展示

#### Scenario: 窗口扩宽恢复显示
- **WHEN** 用户拖拽窗口边缘使宽度 >= 700px
- **THEN** 侧边栏恢复显示,表格恢复分类分组视图

#### Scenario: 窄屏搜索
- **WHEN** 窗口宽度 < 700px 且用户在搜索框输入关键词
- **THEN** 系统过滤匹配的表格,仍以扁平列表展示结果

### Requirement: Flat table list in narrow mode
窄屏扁平列表 SHALL 保留每行的核心信息但不显示多列 grid。

#### Scenario: 扁平列表行内容
- **WHEN** 窗口宽度 < 700px
- **THEN** 每行显示: 文件名 + 行数×列数 + [打开] + [比对]

#### Scenario: 扁平列表不分页
- **WHEN** 侧边栏隐藏
- **THEN** 表格列表连续滚动,无分类 header 分隔
