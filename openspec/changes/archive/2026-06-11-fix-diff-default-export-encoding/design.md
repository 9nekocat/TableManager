## Context

当前 [比对] 按钮调用 `quickDiff()` → `GET /api/diff/smart/<path>` → 应用内 `renderDiffResults()` 展示文字差异。用户反馈默认想在 Excel 里看带颜色的差异 xlsx。

SVN 提交面板文件名乱码: `svn status` 在中文 Windows 输出 GBK 编码,但 `subprocess.run` 强制 `encoding="utf-8"` 解码,非法的 GBK 字节被替换为 �。

## Goals / Non-Goals

**Goals:**
- [比对] 按钮点击默认导出差异 xlsx 到 Excel
- 行内提供下拉菜单切换为应用内查看
- SVN 输出编码自适应系统 locale

**Non-Goals:**
- 不改变右上角"版本比对"按钮的行为
- 不改变合并预览等已有功能

## Decisions

### D1: [比对] 按钮 = 点击直接导出,右侧下拉箭头选应用内

```
[打开] [比对 ▾]
         ├─ Excel 查看 (默认, 直接点击)
         └─ 应用内查看
```

点击按钮主体 → 执行 `smartExportDiff()` (调用 export API)
点击下拉箭头 → 展开菜单,可选择 "应用内查看" 调用原 `quickDiff()`

### D2: 新增 `smart_export_diff()` in diff_engine.py + API route

```
GET /api/diff/smart-export/<path>
→ smart_export_diff(file):
   1. svn status 判断 M vs clean
   2. 调用 export_diff_to_excel(file, revision)
   3. os.startfile(output) 打开 Excel
   4. return {ok, output}
```

### D3: SVN 编码修复

`sencodingsv_codec = locale.getencoding()` 获取系统默认编码 (中文 Windows → `gbk`)
所有 `encoding="utf-8"` → `encoding=locale.getencoding()`

保留 `errors="replace"` 防止单字节损坏导致整个命令失败。

## Risks

- [Risk] `locale.getencoding()` 在部分环境可能返回 `utf-8` 但 SVN 输出是 GBK → [Mitigation] 保留 `errors="replace"`, bad chars still replaced but much less common with correct codec
