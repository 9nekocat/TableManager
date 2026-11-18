# TableManager

面向 UE DataTable 配置表的浏览器、校验、差异比对与三路合并工具。支持 Windows Web GUI、CLI 和 AI 可消费的文本/JSON 输出。

## 主要功能

- 浏览、搜索和校验 `DataTable/Raw` 下的 Excel 配置表。
- 检查重复表头、重复 RowName、类型错误、Excel 错误值和公式风险。
- 两文件差异：先转纯文本快照，公式直接使用 Excel 缓存值。
- 三路合并：支持单元格、行、列的新增、删除和修改。
- 冲突策略：`fail`、`ours`、`theirs`，默认 `fail`，不会静默覆盖。
- 输出差异/合并 Excel，并附带 Conflicts、Changes、Issues 审查页。
- 清理合并结果中的全空行和全空列。
- 保留简单公式，并在行、列移动后重绑定行内引用。
- 非公式但以 `=` 开头的配置字符串使用 Excel quote-prefix 文本格式写回。

## 目录

```text
src/             Python 应用、旧版合并链路和纯数据核心
rust-core/       Rust 校验、差异和三路合并核心
gui/             React 前端
data/            分类和引用检查配置
tests/           Python 单元测试
rust-core/tests/ Rust 单元测试
docs/            重构审查与性能报告
```

## 环境要求

- Windows
- Python 3.11+
- Node.js 18+
- Microsoft Excel 或 WPS
- Rust stable（构建 `rust-core` 时）
- SVN 命令行工具（仅 SVN 更新、锁和冲突工作流需要）

安装依赖：

```powershell
pip install -r requirements.txt
cd gui
npm install
cd ..
```

## 启动

### Web 界面

```powershell
python src/table_server.py --web
```

原生窗口模式：

```powershell
python src/table_server.py
```

前端开发：

```powershell
cd gui
npm run dev
```

构建前端：

```powershell
cd gui
npm run build
cd ..
```

### Rust 核心

```powershell
cargo build --release --manifest-path rust-core/Cargo.toml
```

```powershell
# 校验
rust-core\target\release\table-core.exe validate A:\DataTable\Raw\Effect.xlsx --json

# 差异
rust-core\target\release\table-core.exe diff old.xlsx new.xlsx --format text
rust-core\target\release\table-core.exe diff old.xlsx new.xlsx --format excel --output diff.xlsx

# 三路合并
rust-core\target\release\table-core.exe merge base.xlsx ours.xlsx theirs.xlsx --strategy fail --output merged.xlsx
```

`--strategy fail` 遇到冲突不会写盘；需要查看冲突预览时可加 `--allow-conflicts`。使用 `ours` 或 `theirs` 时，冲突会按指定一侧解决并写入审查报告。

### SVN 比对与合并

仓库内的 `rust-core/svn` 提供可直接配置到 SVN 命令行和 TortoiseSVN 的包装脚本。`.xlsx`、`.xlsm`、`.xlsb`、`.xls` 使用 Rust 表核心，其他文件沿用 xlCompare。

```ini
[helpers]
diff-cmd = D:/Project/TableManager-refactor/rust-core/svn/svn-diff.cmd
diff-extensions =
merge-tool-cmd = D:/Project/TableManager-refactor/rust-core/svn/svn-merge.cmd

[miscellany]
diff-ignore-content-type = yes
```

完整配置方式见 [rust-core/svn/README.md](rust-core/svn/README.md)。

## 配置

### 项目根目录

工具会从自身目录向上查找包含 `DataTable/Raw` 或 `.uproject` 的项目根目录。也可以显式设置：

```powershell
$env:CATSINTHESHELL_ROOT = "D:\catsintheshell"
```

### 数据目录

```text
DataTable/Raw/           Excel 源表
DataTable/Processed/     UE 导出的数据
data/categories.json     表格分类
data/references.json     跨表引用规则
data/effect_checks.json  Effect 参数检查规则
```

临时差异和合并产物输出到调用方指定路径；旧工作流默认使用 `DataTable/Raw/.temp/`。

## Excel 约定

- 只有 `#` 开头的 sheet 作为数据页。
- 第 1 行：中文字段名。
- 第 2 行：英文字段名。
- 第 3 行：类型。
- 数据从第 4 行开始。
- A 列为 RowName，行匹配使用 `(RowName, 重复序号)`，不会静默覆盖重复行。

## 公式规则

支持：

- 行内单元格或范围引用。
- 跨页固定目标引用，例如 `'#Lookup'!$A$1`。

会提示风险：

- 跨行相对引用。
- 跨页非固定引用。
- 跨工作簿引用。
- 动态数组、易失函数和计算上下文敏感函数。

合并输出会保留简单公式；复杂外部链接、动态数组元数据和复杂格式的原地写回仍使用原有 Python/Excel COM 安全链路。

## 测试

```powershell
python -m pytest tests -q
cargo test --manifest-path rust-core/Cargo.toml
```

## 更多文档

- [核心重构审查](docs/core-refactor-review.md)
- [真实配表 Benchmark 与 Rust 实测](docs/real-table-benchmark-rust-estimate.md)
- [Rust 核心说明](rust-core/README.md)
