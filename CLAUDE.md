# CLAUDE.md

本文件为 Claude Code (claude.ai/code) 在此仓库中工作时提供指导。

## 概述

`catsintheshell` 游戏项目的 UE DataTable 配置表管理工具。用于浏览、校验和合并 `DataTable/Raw/` 下的 Excel 配置表。提供 Web GUI（Flask + React/pywebview）和 CLI 两种使用方式。

## 构建与运行

```bash
# 安装依赖
pip install -r requirements.txt
cd gui && npm install && cd ..

# 开发：Web 服务器（浏览器模式）
python src/table_server.py --web

# 开发：原生窗口（需要 pywebview）
python src/table_server.py

# CLI 模式
python src/table_manager.py list
python src/table_manager.py check

# 构建 React 前端
cd gui && npm run build && cd ..

# 打包为独立 exe（先构建 diff.exe，TableManager.spec 的 datas 依赖它）
pyinstaller diff_standalone.spec
pyinstaller TableManager.spec
```

## 运行测试

```bash
# 单元测试（pytest 或直接运行）
python -m pytest tests/ -v

# 合并集成测试（13 个测试用例，含 xlsx 测试数据）
python src/run_tests.py
# 或
python src/test_runner.py
```

测试需要安装 Excel（通过 COM 重算公式单元格后再比对）。没有 Excel 时会回退到公式字符串匹配，但基于值的 key 匹配可能不够精确。

## 目录结构

```
src/                     全部 Python 代码（模块间互相 import，入口文件自带 src 路径引导）
  table_server.py          Web 服务入口（Flask + pywebview 原生窗口）
  table_manager.py         CLI 入口（argparse 子命令）
  manager_core.py          核心逻辑：文件发现、校验、分类、项目根目录定位；DATA_DIR 指向 data/
  merge_engine.py          Excel 三路合并（冲突检测 → 预览 → 应用）
  addmerge_engine.py       增量合并（staging → Raw 目标表：只增只改不删行，CLI `add-merge`）
  merge_workflow.py        事务化合并写回（备份 → 应用预览 → 失败回滚），merge resolve 走这里
  diff_engine.py           SVN 版本差异比对（当前版本 vs 任意 SVN 版本）
  svn_client.py            SVN 子进程封装（status、update、commit、lock/unlock）
  formula_utils.py         Excel 公式解析、显示格式化、引用调整
  excel_recalc.py          Excel COM 自动化（重算公式、解析外部链接）
  diff_standalone.py       TortoiseSVN 外部差异比对工具入口
  ref_index.py             跨表引用索引（按 references.json + effect_checks.json 规则扫描 → 反向索引 + 悬空检测）
  table_checks.py          导表检查规则集（表头/RowName/类型/引用完整性，error/warning 分级）
  processed_checks.py      Processed 导出数据（DT_*.json）引用检查（导出后精确分析，低噪音）
  scan_categories.py       categories.json 校准工具（扫描 Raw 实际文件，报告/修复缺失与漏编目）
data/                     运行时配置（categories.json / references.json / effect_checks.json / processed_checks.json）
tests/                    单元测试（pytest）；tests/diff用例/ 为手工 diff 素材
测试用例/                 合并集成测试数据（run_tests.py / test_runner.py 使用）
ref/                      参考文档（脚本配置方法.xlsx）
gui/                      前端源码与构建产物（gui/dist 打包进 exe）
art/ env/ openspec/       图标 / SVN 安装包 / OpenSpec 规格
```

配置路径解析（`manager_core._get_data_dir()`）：源码运行时 = `src/` 的上级 `data/`；PyInstaller 打包后 = `_MEIPASS/data`。前端产物 `gui/dist` 同理（源码 = 项目根 `gui/dist`，打包后 = `_MEIPASS/gui/dist`）。

### 前端 (gui/)

React + Vite 单页应用，位于 `gui/src/`。组件：`App.jsx`（根组件）、`Sidebar.jsx`（分类导航）、`TableGrid.jsx`/`TableRow.jsx`（主表格列表）、`Toolbar.jsx`、`MergePanel.jsx`、`DiffPanel.jsx`、`CommitPanel.jsx`、`ValidationPanel.jsx`（导表检查）、`RefsPanel.jsx`（引用查找）、`Toast.jsx`、`ParticleBg.jsx`。状态管理通过 React Context（`context/AppContext.jsx`）。构建产物 `gui/dist/` 作为静态文件打包进 PyInstaller exe。

### 项目根目录检测

`manager_core.py` 中的 `_find_project_root()` 从 exe/脚本所在目录向上查找包含 `DataTable/Raw` 或 `*.uproject` 的目录。回退到 `CATSINTHESHELL_ROOT` 环境变量。所有文件路径均相对于 `RAW_DIR` = `PROJECT_ROOT/DataTable/Raw`。

### Excel 数据表约定

- 仅 `#` 开头的 sheet 是数据 sheet（如 `#1`、`#CardData`）
- 第 1 行：中文字段名，第 2 行：英文字段名，第 3 行：类型信息
- 数据行从第 4 行开始（`DATA_START_ROW = 4`）
- A 列 = RowName（身份标识 key，可能是公式如 `=B4`）
- `categories.json` 将相对文件路径映射到分类，由 `load_categories()` 加载

### 合并引擎流程

1. `ConflictDetector` — 对所有 `#` 开头的 sheet 进行三路差异比对（base/mine/theirs）。每个文件读取两遍：`data_only=True` 取值，`data_only=False` 取公式。
2. `MergePreview` — 生成包含 `[我的]`/`[仓库]`/`[最终]` 行的预览 xlsx。黄色=已修改，绿色=决策行，红色=冲突单元格。A 列=[标记]，B 列=[修改类型] 下拉框，C 列起=数据。
3. `MergeResolver.resolve()` — 从预览文件中读取 `[最终]` 行的决策，应用到 mine 文件（作为非冲突数据的基底），通过 `formula_utils.adjust_formula()` 调整公式引用。

合并核心行为：mine 文件作为基底工作簿 — 所有非冲突行、sheet、公式完整保留。仅冲突行使用 `[最终]` 行的值。新增行通过与 theirs 文件的邻居匹配，插入到正确位置；但当键来自计算值（`keys_from_values`）且 `[最终]` 行的 colA(RowName) 已存在于 mine 时，一律视为对既有行的修改（如 theirs 改了 B 列注释），绝不作为新行插入 — 否则会产生重复 RowName。外部链接通过在 openpyxl 保存后注入 zip 级 XML 来保留。

### 公式处理

- `formula_utils.format_display()`：显示为 `'=FORMULA → VALUE`（单引号前缀，使 Excel 将其视为文本）
- `formula_utils.adjust_formula()`：合并过程中插入/删除行后，重新映射公式中的行号
- `formula_utils.parse_cell_refs()`：提取单元格/范围引用，包括跨 sheet（`Sheet2!B4`）和跨文件（`[file.xlsx]Sheet!A1`）
- `merge_engine` 中的 `_strip_col()`：在 resolve 阶段将显示格式标准化回 VALUE，用于 key 匹配

### 外部链接保留

openpyxl 在保存时会损坏外部链接 XML。`_read_external_link_cache()` 在 openpyxl 读取之前在 zip 层面缓存原始的外部链接数据，`_inject_external_links()` 在保存后恢复。

### 数组公式与动态数组元数据保留

openpyxl 保存时会丢失 `xl/metadata.xml`（XLDAPR 动态数组元数据）和单元格的 `cm` 属性，导致含 XLOOKUP 等新函数的数组公式在 Excel 重算时退化为隐式交集（#VALUE! 或只取首元素）。`_inject_dynamic_array_metadata()` 在 resolve 保存后于 zip 层面恢复 metadata.xml 并给所有 `<f t="array">` 单元格补 `cm="1"`（与 cm 单元格一一对应，同外部链接注入模式）。读取端 `_read_all_rows_with_formulas` 会把 openpyxl 的 `ArrayFormula` 对象解包为公式文本进入合并管线；resolve 写回时对原本是数组公式的单元格重新包装为 `ArrayFormula` 以保留 CSE 语义。

### 导表检查与引用查找

- `table_checks.py` 的规则在 `manager_core.validate_table()` 既有检查（Excel 错误值、重复 RowName）之后追加执行，错误 dict 结构相同，额外带 `level`（error/warning）与 `message`。规则：英文字段名重复、RowName 为空、类型不匹配（按第 3 行类型）、ID/公式引用不存在、配置文件缺失为 error 级；中文字段名重复、表头缺失、空数据行为 warning 级。
- `ref_index.py` 按 `references.json` + `effect_checks.json` 扫描全部配置表建立反向引用索引（首次约 20 秒，按 mtime 签名缓存）。引用类型：`column`（列级 ID 引用，支持数组拆分与多候选目标）、`text`（描述文本中的 `{Buff,1001}` 式内嵌引用）、`effect_params`/`effect_condition`（Effect.xlsx 按 Type 的 Params 位置语义）、`formula`（仅跨 sheet/外链公式，表内拼接公式与指向辅助 sheet 的公式视为噪音跳过）。
- 目标 ID 空间默认取目标表 A 列 RowName（合并该文件所有 `#` sheet），也可用 `target.column` 指定其他列（如 Tag 引用 GameplayTagTableRow 的 Tag 列）。
- **性能注意**：openpyxl `read_only=True` 模式下 `ws.cell()` 随机访问会重复流式解析（大表 O(n²)），`ref_index`/`table_checks` 均使用 eager 加载。
- 入口：CLI `python src/table_manager.py refs <表> [--sheet S] [--row R] [--dangling]`；API `GET /api/refs`、`GET /api/refs/dangling`；GUI `RefsPanel`（工具栏"引用"按钮）。导表检查沿用 `check` 子命令与 `/api/check`。
- `Effect.xlsx` 的 `Params` 列是位置参数（含义由 `Type` 决定），规则写在 `effect_checks.json`（语义一手来源：小闪在线文档[脚本配置方法]）。`pos` 支持数字或 `"N+"`；`kind` 为 `ref` 的位置做悬空检查，`literal` 仅作文档。未配置 `ref` 规则的 Type/位置不检查。条件列 `ConditionIDs` 的目标由全局 `condition_targets` 指定。配置方法见 `.agents/skills/effect-check-config/SKILL.md`。
- `scan_categories.py` 用于校准 categories.json：`python src/scan_categories.py` 报告配置中存在但磁盘缺失、磁盘存在但未编目的文件；`--write` 自动删缺失、把新文件加入 uncategorized（先备份）。

### 导出数据检查（Processed 模式）

xlsx 模式（check/refs）做导表前检查；`processed_checks.py` 做**导出后精确分析**，数据源为 `DataTable/Processed/DT_*.json`（UE DataTable 导出产物）：

- 导出数据天然去噪：无公式/辅助 sheet/垃圾行，Type 列为自描述枚举名（`ECardEffectType::Event` 等），位置参数规则按枚举名配置（`processed_checks.json` 的 `positional_rules`），语义与 `effect_checks.json` 的数字 Type 一一对应
- 编码混杂（utf-8-sig / utf-16）自动探测；字段值兼容原生 list、字符串 repr、逗号串
- `'0'` 是各表通用的"无"占位符，检查中跳过（但不能在 token 化时删除——会破坏位置参数的位置）
- 已知导出覆盖缺口：Dialog 仅导出 `#main`、GameplayTag 导出为空、部分表缺字段——相关规则不配置或由引擎自动跳过（报告在 skipped 中）
- `Processed/*.csv` 是陈旧测试数据，一律忽略
- 入口：CLI `python src/table_manager.py pcheck [--table X] [--csv out]`；API `GET /api/pcheck`；GUI ValidationPanel 的"导出数据检查"切换

### 临时目录

所有产物（合并预览、差异导出）存放在 `RAW_DIR/.temp/`（别名为 `TEMP_DIR`）。该目录位于工作副本内，SVN 会自动忽略（或应加入 svn:ignore）。

### PyInstaller 打包

`TableManager.spec` 打包内容：`src/table_server.py` 为入口，`gui/dist/` 为静态文件，`data/`（4 个 json 配置）、`src/diff_standalone.py`、`dist/diff.exe`、`art/` 目录。隐藏导入包括 `win32com`、`webview`、`clr`，用于 Windows COM 和原生窗口支持。

## 关键约束

- 仅限 Windows（COM 自动化、pywebview、`os.startfile`、`win32com`）
- 需要安装 Microsoft Excel 用于公式重算和打开文件
- 版本控制功能需要 SVN 命令行工具（`svn.exe`）— 通过注册表和常见路径自动检测
- 中文 Windows 上 SVN 输出为 GBK 编码；`_decode_svn()` 处理多编码回退
- `recalc_xlsx()` 期间，外部链接工作簿必须在其原始路径上可访问
