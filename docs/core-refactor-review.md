# 合并与比对核心重构审查

## 范围与基线

- 分支：`refactor/core-diff-merge`
- 审查对象：`merge_engine.py`、`diff_engine.py`、`addmerge_engine.py`、`formula_utils.py`、`table_checks.py` 及其现有测试。
- 旧测试基线：设置 `CATSINTHESHELL_ROOT` 后，原有 130 个 pytest 测试全部通过。运行日志仍出现 Excel COM 的 `0x80010108` 异常提示，但测试最终通过，说明 COM 路径对运行环境高度敏感。
- 新核心测试：`tests/test_core_refactor.py`。

## 主要问题

### P0：重复 RowName 会被静默覆盖

`merge_engine.py:498-499` 检测到重复 RowName 后只打印警告，随后用字典赋值覆盖上一行。后续差异、预览和合并只看到最后一行，属于数据丢失风险。

新核心保留所有重复行，并把重复 RowName 作为 error 报告；内部使用 `(RowName, 出现序号)` 作为行身份，不再静默覆盖。

### P0：核心逻辑与 UI/SVN/Excel COM 强耦合

现有合并入口 `generate_merge_preview` / `resolve_merge` 直接生成 Excel 预览，合并器同时负责行定位、公式调整、COM 写入、外链恢复和动态数组元数据恢复。核心能力无法被 AI、CLI 或其他界面稳定复用。

新核心先建立纯数据 API：读表、校验、CRUD、两路差异、三路合并和报告输出均不依赖 GUI、SVN 或 Excel COM。

### P1：差异应直接比较公式计算值

`diff_engine.py:241-244` 和 `_compare_data` 比较缓存值，但旧流程仍把 workbook 对象作为比较中间层，并依赖 Excel COM 预先重算。

新核心按重构要求先转为纯文本快照，再比较文本；公式只读取 Excel 缓存值。无缓存时按空值处理并发出警告。公式风险仍由独立校验链分级：

- 行内引用、行内范围：允许。
- 跨页固定目标（行列均使用 `$`）：允许。
- 跨行相对引用、跨页非固定引用、跨工作簿引用：报错。
- 动态数组、易失函数、`XLOOKUP`、`FILTER`、`OFFSET` 等：警告。

### P1：比对速度主要被 Excel 重算阻塞

`diff_engine.py:642` 在双文件比较前调用 `recalc_batch`。这要求安装 Excel，启动 COM 会话，并复制两份文件；表越大、外链越多，延迟和异常概率越高。

新核心不启动 Excel，先把表冻结成纯文本，再执行线性文本比较，避免旧链路中的 COM 往返与公式 AST 比较。

A 盘真实表基准与 Rust 预估见 `docs/real-table-benchmark-rust-estimate.md`。

### P1：表头重复不是统一主流程的硬性结果

`table_checks.py:66` 已有重复表头检查，但差异/合并主流程本身没有统一的警告结果。重复英文名会产生语义歧义，重复完整表头也会影响列对齐。

新核心在同一份校验报告中统一输出：

- 重复英文字段名：error。
- 重复中文字段名：warning。
- 重复完整表头：warning。
- 有数据但缺英文字段名：warning。

### P1：空行/空列清理不完整

旧合并主要处理尾部行清理和连续行重排，没有统一的“输出工作簿无全空行/全空列”契约。

新核心在加载时移除全空数据行并记录警告，在保存前再次清理全空行和全空列。清理只针对整行/整列全空，不会误删正常业务中的局部空单元格。

### P1：缺少 AI 可直接消费的稳定结果

旧结果主要面向 Excel 预览，内部 dict 没有稳定的文本契约。

新核心提供：

- 结构化 JSON。
- 类 Git 的统一文本差异。
- 可阅读的 Excel 汇总、差异、冲突和警告页。

## 新核心边界

| 模块 | 职责 |
| --- | --- |
| `src/core/model.py` | 工作表模型、加载/保存、CRUD、表头与空行空列清理、校验 |
| `src/core/formula.py` | 简单公式风险分级、行内公式重绑定 |
| `src/core/text.py` | 把工作簿冻结为纯文本快照，公式只取缓存值 |
| `src/core/diff.py` | 文本快照差异、行列增删改、文本/JSON/Excel 输出 |
| `src/core/merge.py` | 三路合并、冲突策略、审查报告 |
| `src/core/operations.py` | 面向人类和 AI 的声明式 JSON 操作 |
| `src/core_cli.py` | `validate`、`inspect`、`diff`、`merge`、`edit` |

## 使用入口

```powershell
# 校验与查询
python src/core_cli.py validate Effect.xlsx --json --clean
python src/core_cli.py inspect Effect.xlsx --sheet "#Main" --json

# 两文件比对
python src/core_cli.py diff before.xlsx after.xlsx --format text
python src/core_cli.py diff before.xlsx after.xlsx --format json
python src/core_cli.py diff before.xlsx after.xlsx --format excel --output diff.xlsx

# 三路合并；默认 fail，不会静默选择冲突
python src/core_cli.py merge base.xlsx mine.xlsx theirs.xlsx --output merged.xlsx
python src/core_cli.py merge base.xlsx mine.xlsx theirs.xlsx --strategy ours --output merged.xlsx
```

`edit` 接收有序 JSON 操作数组，支持 `set_cell`、`add_row`、`delete_row`、`add_column`、`update_column`、`delete_column`。位置参数为零基索引。

```json
[
  {
    "op": "set_cell",
    "sheet": "#Main",
    "row": "Card_001",
    "column": "Power",
    "value": 12
  },
  {
    "op": "add_column",
    "sheet": "#Main",
    "position": 2,
    "header": {"cn": "备注", "en": "Note", "type": "string"},
    "default": ""
  }
]
```

## 当前验收项

- 单元格、行、列：增、删、改、查。
- 重复表头：显示警告/错误，不合并列。
- 重复 RowName：全部保留并报错。
- 合并、输出差值、输出合并结果：清理全空行、全空列。
- 简单公式：可识别、可在行列移动时重绑定。
- 复杂公式：显式提示风险，不声称完全兼容。
- 人类输出：Excel；AI 输出：JSON、类 Git 文本。

## 明确不做或暂缓

- 本轮不替换旧 GUI 和旧 CLI 入口，先验证新核心 API。
- 新核心的生成文件不承诺保留 Excel COM 缓存、宏、外链 XML 或动态数组元数据。
- 本轮不引入 Rust。当前环境没有 Rust 工具链；旧链路最明显的性能瓶颈是 Excel COM，新核心已移除该依赖。只有在真实业务表基准显示 Python 数据层仍是热点时，才值得把 `diff.py` 的纯计算部分迁移到 Rust，而不是重写整个程序。
- 列删除若被公式引用，默认拒绝；确认破坏风险后才允许执行。

## 下一阶段

1. 用真实大表建立加载、diff、merge 基准和内存峰值基线。
2. 将旧 GUI/CLI 的差异与合并按钮适配到新核心，保留兼容入口。
3. 增加真实 SVN 三文件样本回归，覆盖列重命名、重复表头、外链公式和动态数组。
4. 根据基准决定是否需要 Rust 差异内核；优先替换热点，而不是整体重写。
