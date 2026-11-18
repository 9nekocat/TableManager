## ADDED Requirements

### Requirement: Always-on-top toggle
系统 SHALL 提供大头钉按钮切换窗口置顶状态。

#### Scenario: 激活置顶
- **WHEN** 用户点击 toolbar 图钉按钮 (当前未激活)
- **THEN** 窗口设为 `HWND_TOPMOST`,按钮变为激活样式 (filled 蓝色)

#### Scenario: 取消置顶
- **WHEN** 用户再次点击图钉按钮 (当前已激活)
- **THEN** 窗口取消 `HWND_TOPMOST`,恢复普通 Z-order,按钮变为默认样式 (outline)

#### Scenario: 置顶状态持久化
- **WHEN** 应用重启
- **THEN** 置顶状态不保留,默认未激活

### Requirement: Pin button in toolbar
图钉按钮 MUST 位于 toolbar 左侧,搜索框之后。

#### Scenario: 图钉按钮可见性
- **WHEN** 窗口正常显示
- **THEN** toolbar 中显示图钉按钮 (📌),tooltip "窗口置顶"
