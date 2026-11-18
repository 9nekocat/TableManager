## Why

1. 行内 [比对] 按钮默认做应用内 diff 展示,但用户实际最常用的场景是: 点比对→直接在 Excel 中看到差异(带颜色标记的 xlsx)。应用内文字 diff 作为辅助即可。
2. SVN 提交面板中部分文件名显示为乱码 — SVN 输出的 GBK 编码被强制 UTF-8 解码导致。

## What Changes

- **[比对] 按钮行为改为默认导出到 Excel**: 点击 [比对] 直接调用 `POST /api/diff/export`,生成差异 xlsx 并在 Excel 中打开
- **[比对] 按钮旁加下拉/右键菜单**: 提供 "在 Excel 中查看" (默认) 和 "应用内查看" 两个选项
- **SVN 编码修复**: `svn_client.py` 中所有 `subprocess.run` 的 `encoding` 从硬编码 `utf-8` 改为 `locale.getencoding()` 自动取系统编码

## Capabilities

### New Capabilities

(无——修改现有行为)

### Modified Capabilities

- `inline-smart-diff`: [比对] 按钮默认行为改为导出 Excel; 应用内查看变为可选的附加项

## Impact

| 层 | 影响 |
|---|------|
| `templates/index.html` | `quickDiff()` 改为默认调用 export API; 加下拉菜单 |
| `svn_client.py` | 所有 `encoding="utf-8"` → `encoding=None` (自动) |
| `diff_engine.py` | `smart_diff()` 可能需要新增 `smart_export_diff()` 变体 |
| `table_server.py` | 可能新增 `GET /api/diff/smart-export/<path>` |
