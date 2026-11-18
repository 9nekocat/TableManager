## Why

点击 [比对] 按钮后, diff 生成(含 Excel 重算)可能需要 5-15 秒。当前无加载提示,用户不确定是否正在处理。

## What Changes

- **[比对] 导出 Excel**: 按钮显示 spinner + "比对中..." 文字, 完成后 toast 提示
- **[比对] 应用内查看**: diff-panel 显示 spinner + "正在比对 <filename>..." 
- **手动版本比对**: 点击"比对"按钮后显示 spinner + 进度文字

## Capabilities

### Modified Capabilities

- `inline-smart-diff`: 比对按钮增加加载态反馈

## Impact

| 文件 | 改动 |
|------|------|
| `gui/src/components/TableRow.jsx` | `quickDiff()` / `quickDiffInApp()` 加 loading state |
| `gui/src/components/DiffPanel.jsx` | 手动比对按钮加 spinner |
