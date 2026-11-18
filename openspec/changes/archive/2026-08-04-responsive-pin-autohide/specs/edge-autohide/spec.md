## ADDED Requirements

### Requirement: Edge auto-hide when window docked to screen edge
当窗口贴靠屏幕左/右边缘时,系统 SHALL 在鼠标移出后自动将窗口滑出屏幕(仅留 4px 抓手),鼠标移入抓手区域时滑回原位。

#### Scenario: 贴左边缘 — 鼠标移出隐藏
- **WHEN** 窗口左边 ≤ 0px 且鼠标向右移出窗口超过 20px
- **THEN** 窗口水平滑出,仅留右侧 4px 可见

#### Scenario: 贴左边缘 — 鼠标移入展开
- **WHEN** 窗口处于隐藏状态且鼠标移动到屏幕左侧 8px 范围内
- **THEN** 窗口滑回原位 (left=0)

#### Scenario: 贴右边缘 — 鼠标移出隐藏
- **WHEN** 窗口右边 ≥ 屏幕宽度且鼠标向左移出窗口超过 20px
- **THEN** 窗口水平滑出,仅留左侧 4px 可见

#### Scenario: 贴右边缘 — 鼠标移入展开
- **WHEN** 窗口处于隐藏状态且鼠标移动到屏幕右侧 8px 范围内
- **THEN** 窗口滑回原位

#### Scenario: 窗口不贴边
- **WHEN** 窗口不靠屏幕左/右边缘 (left > 20px 且 right < screenWidth - 20px)
- **THEN** autohide 逻辑不触发,窗口保持原位

### Requirement: Autohide toggle
Autohide 功能 MUST 可开关,默认关闭。

#### Scenario: 开启 autohide
- **WHEN** 用户在 toolbar 点击 autohide 按钮
- **THEN** 启动位置检测线程,按钮变为激活样式

#### Scenario: 关闭 autohide
- **WHEN** 用户在 toolbar 再次点击 autohide 按钮
- **THEN** 停止检测线程,如果窗口处于隐藏状态则展开,按钮恢复默认样式

### Requirement: Autohide performance
检测线程 MUST NOT 造成明显 CPU 占用。

#### Scenario: 轮询间隔
- **WHEN** autohide 启用
- **THEN** 窗口位置检测间隔为 200ms,线程优先级为最低
