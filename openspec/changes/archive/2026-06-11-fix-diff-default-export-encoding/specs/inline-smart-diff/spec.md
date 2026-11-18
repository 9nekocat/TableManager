## MODIFIED Requirements

### Requirement: Inline smart diff button on table rows
每个表格行 SHALL 显示 [比对] 按钮,点击默认导出差异 xlsx 并在 Excel 中打开。按钮右侧下拉箭头提供"应用内查看"选项。

#### Scenario: 默认点击 — 导出到 Excel
- **WHEN** 用户点击表格行 [比对] 按钮主体
- **THEN** 系统调用 `GET /api/diff/smart-export/<path>`,生成差异 xlsx,在 Excel 中打开,并显示 toast "比对结果已在 Excel 中打开"

#### Scenario: 下拉菜单 — 应用内查看
- **WHEN** 用户点击 [比对] 右侧 ▾ 并选择 "应用内查看"
- **THEN** 系统调用 `GET /api/diff/smart/<path>`,底部 diff-panel 展开显示文字差异 (原行为)

#### Scenario: 已修改文件导出 vs HEAD
- **WHEN** 用户点击导出且文件 SVN 状态为 M
- **THEN** 系统 diff vs HEAD,导出 xlsx,打开 Excel

#### Scenario: 未修改文件导出 vs PREV
- **WHEN** 用户点击导出且文件未修改
- **THEN** 系统 diff vs PREV,导出 xlsx,打开 Excel

#### Scenario: 无历史版本
- **WHEN** 文件无历史版本可供 diff
- **THEN** 系统显示 toast 提示 "该文件无历史版本"

### Requirement: SVN output encoding
SVN 子进程输出 MUST 使用系统默认编码解码,而非硬编码 UTF-8,以支持中文文件名。

#### Scenario: 中文文件名在提交面板中正确显示
- **WHEN** SVN 返回包含中文文件名的 status 输出
- **THEN** 提交面板中文件名正确显示中文,无乱码

#### Scenario: 编码自适应
- **WHEN** SVN 输出编码与系统 locale 一致 (中文 Windows = GBK)
- **THEN** `subprocess.run` 使用 `locale.getencoding()` 解码,并保留 `errors="replace"` 防损坏

## ADDED Requirements

### Requirement: Standalone diff tool packaging
diff_standalone.py 及相关依赖 MUST 打包进 TableManager.exe,支持智能比对导出功能。

#### Scenario: 导出 diff 可独立使用
- **WHEN** 智能比对导出 API 被调用
- **THEN** 系统使用打包内的 diff_standalone 逻辑生成差异 xlsx,不依赖外部 Python 环境
