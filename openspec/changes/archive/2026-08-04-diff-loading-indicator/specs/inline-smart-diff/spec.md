## MODIFIED Requirements

### Requirement: Inline smart diff button loading state
点击 [比对] 按钮时,按钮 MUST 显示加载状态直到 API 返回或超时。

#### Scenario: 导出模式加载
- **WHEN** 用户点击 [比对] 按钮主体
- **THEN** 按钮文字变为 spinner + "比对中...", 按钮 disabled, 直到 API 返回后恢复

#### Scenario: 应用内查看加载
- **WHEN** 用户点击 "应用内查看" 
- **THEN** diff-panel 展开并显示 "正在比对 <filename>..." + spinner, API 返回后渲染结果

#### Scenario: 手动比对加载
- **WHEN** 用户在版本比对面板选择文件+版本并点击"比对"
- **THEN** 按钮显示 spinner + "比对中...", 面板显示 loading, API 返回后恢复
