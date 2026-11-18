## 1. [比对] 按钮改为默认导出 Excel

- [x] 1.1 新增 `GET /api/diff/smart-export/<path>` 路由: smart 判断版本 → `export_diff_to_excel()` → `os.startfile()` 打开
- [x] 1.2 前端: `quickDiff()` 改为默认调用 `smart-export` API, toast 提示 "比对结果已在 Excel 中打开"
- [x] 1.3 [比对] 按钮改为 split-button 样式: 主体点导出, 右侧 ▾ 下拉菜单可选 "应用内查看"
- [x] 1.4 下拉菜单 CSS + JS: 点击下拉箭头展开菜单, 点击外部关闭
- [x] 1.5 "应用内查看" 选项调用原 `/api/diff/smart/<path>` 逻辑

## 2. SVN 编码乱码修复

- [x] 2.1 `svn_client.py`: 顶部加 `import locale`, 定义 `_SVN_ENCODING = locale.getencoding()`
- [x] 2.2 所有 `subprocess.run` 的 `encoding="utf-8"` 替换为 `encoding=_SVN_ENCODING`
- [x] 2.3 保留 `errors="replace"` 防止单字节损坏

## 3. diff 工具打包

- [x] 3.1 `TableManager.spec` datas 加 `('diff_standalone.py', '.')` 确保打包进 EXE
- [x] 3.2 确认 `diff_engine.py` 和 `excel_recalc.py` 已在隐藏导入中

## 4. 验证

- [x] 4.1 测试 [比对] 按钮默认导出 Excel 并打开
- [x] 4.2 测试下拉菜单 "应用内查看"
- [x] 4.3 测试 SVN 提交面板中文文件名无乱码
- [x] 4.4 PyInstaller 打包并验证 EXE
