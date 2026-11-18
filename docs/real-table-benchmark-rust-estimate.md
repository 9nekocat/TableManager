# A 盘真实配表 Benchmark 与 Rust 实测

## 结论

Rust 只读核心已经落地到 `rust-core/`，采用“xlsx → 纯文本快照 → 校验/差异”流程，公式只取 Excel 缓存值。

在 `A:/DataTable/Raw` 的真实表上：

- 两个大表的端到端差异分别约 **4.28 秒（Card）** 和 **2.07 秒（Effect）**。
- 纯文本差异本身约 **0.09 秒**，已不是瓶颈。
- 主要成本是 xlsx 解包和构建模型：大表占用约 **80%**。
- 全目录 36 个有效 xlsx 的一遍加载、校验、自差异：Python 约 **5.71 秒**。

实际编译后的 Rust release 核心：

- Card 两文件差异：`4.285s → 0.394s`，约 **10.9 倍**。
- Effect 两文件差异：`2.065s → 0.334s`，约 **6.2 倍**。
- 全目录扫描：`5.705s → 0.488s`，约 **11.7 倍**。

实测高于原先的中性预估；主要原因是 calamine 避免了 openpyxl 的完整对象树和外部链接对象构建。

## 测试环境与口径

- 日期：2026-11-18
- 原始目录：`A:/DataTable/Raw`
- 文件：36 个有效 xlsx；另外发现两个 `~$*.xlsx` Excel 锁文件，无法作为 zip 打开
- 轮次：主要表 3 轮，取中位数
- 修改量：每张表 100 个单元格
- 原始表只读；临时副本和输出均放在系统临时目录
- 差异口径：先转为纯文本快照；公式只使用 Excel 缓存值；无缓存时按空值并警告

## Rust 实测

| 场景 | Python | Rust release | 提升 |
| --- | ---: | ---: | ---: |
| Card 两文件差异（CLI） | 4.285s | 0.394s | 10.9x |
| Effect 两文件差异（CLI） | 2.065s | 0.334s | 6.2x |
| 36 表加载+校验+自差异 | 5.705s | 0.488s | 11.7x |

Rust 核心内部分段（Card，100 个单元格修改）：

| 阶段 | 中位时间 |
| --- | ---: |
| xlsx 加载 | 0.104s |
| 文本快照 + 差异 | 0.022s |
| 文本格式化 | 0.00006s |
| 差异 Excel 写出 | 0.002s |

Effect 的文本差异约 0.040s，差异 Excel 写出约 0.002s。

## 代表性大表

| 表 | 数据行 | 单元格 | 公式单元格 | 单次加载 | 纯差异计算 | 含输入校验的差异 | 端到端差异 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `Localization/Card.xlsx` | 2,094 | 123,546 | 66,672 | 1.882s | 0.090s | 1.216s | 4.285s |
| `Effect.xlsx` | 9,083 | 157,560 | 28,837 | 0.830s | 0.088s | 0.468s | 2.065s |

“纯差异计算”包含两份表转文本快照，不包含 xlsx 加载和完整校验。

## 输出成本

| 表 | 文本差异 | 差异 Excel | 三路合并模型 | 合并 Excel 写出 | 端到端合并 |
| --- | ---: | ---: | ---: | ---: | ---: |
| Card | 0.0003s | 0.033s | 0.491s | 1.517s | 5.630s |
| Effect | 0.0001s | 0.024s | 0.666s | 1.872s | 3.428s |

文本输出的时间可以忽略。Excel 输出的时间主要来自工作簿构造和压缩。

## 中小表

| 表 | 单元格 | 单次加载 | 端到端差异 | 端到端合并 |
| --- | ---: | ---: | ---: | ---: |
| `Localization/Buff.xlsx` | 5,888 | 0.284s | 0.625s | 0.916s |
| `ProjectileTable.xlsx` | 1,728 | 0.251s | 0.552s | 0.825s |

即使数据只有几千个单元格，单次加载仍接近 0.25 秒，说明固定 xlsx/openpyxl 解析成本明显。这是 Rust 最值得优化的部分。

## 全目录扫描

| 项目 | 时间 | 占比 |
| --- | ---: | ---: |
| 36 个有效 xlsx 加载 | 4.075s | 71.4% |
| 全表校验 | 1.021s | 17.9% |
| 自差异 | 0.608s | 10.7% |
| 合计 | 5.705s | 100% |

## 当时的 Rust 预估方法

使用测得的阶段占比和 Amdahl 定律：

```text
总加速比 = 1 / Σ(阶段占比 / 阶段加速比)
```

考虑 5% 固定开销（进程启动、文件系统、序列化等）：

| 场景 | xlsx 加载 | 校验 | 文本差异 | 全目录预估 |
| --- | ---: | ---: | ---: | ---: |
| 保守 | 3x | 5x | 5x | 3.0x，约 1.89s |
| 中性 | 5x | 8x | 10x | 4.6x，约 1.24s |
| 乐观 | 8x | 15x | 20x | 6.6x，约 0.86s |

这组阶段加速比基于 Python 剖析结果：

- Python/openpyxl 在构建对象树、解析外部链接、创建每个单元格对象上花费大量时间。
- ZIP 解压和原始文件读取很小；Card 的原始字节读取中位数不足 1ms。
- Rust 若使用 `calamine`/`quick-xml` 直接构建轻量文本模型，避免 openpyxl 的完整对象图和外部链接对象树，加载阶段获得 5 倍左右提升是合理中心值。
- 若 Rust 只是通过 FFI 调用 openpyxl，或者仍构造同样复杂的对象模型，预计只有 1.2–1.8 倍，不值得。

## 已落地的最小 Rust 边界

已实现一个只读 Rust 核心和极简 CLI：

```text
xlsx → 文本快照 → 校验结果 + 差异结果（JSON/类 Git 文本）
```

保留在 Python 的职责：

- Flask/React GUI、pywebview。
- SVN 状态、更新、锁与提交。
- Excel COM、事务备份、回滚。
- 需要精确保留外链、动态数组元数据和复杂格式的写回。

这样收益最大、风险最小。

构建与使用：

```powershell
cargo build --release --manifest-path rust-core/Cargo.toml
rust-core/target/release/table-core.exe validate <file.xlsx> --json
rust-core/target/release/table-core.exe diff <old.xlsx> <new.xlsx> --format text
rust-core/target/release/table-core.exe diff <old.xlsx> <new.xlsx> --format excel --output diff.xlsx
```

## 不应过度承诺的场景

- A 盘冷启动、网络抖动、杀毒软件扫描可能主导时间，Rust 无法消除。
- Excel 文件写出仍受 ZIP 压缩、磁盘和 OOXML 结构约束。
- 复杂公式兼容、动态数组元数据修复、外链保留仍应留在成熟写回链路中。
- 如果最终只优化 Python 中的文本 diff（当前占比约 10%），端到端最多只有约 10% 提升。

## 原始数据

- [代表性表 Benchmark](D:/Project/TableManager-refactor/benchmarks/real_table_benchmark_2026-11-18.json)
- [全目录 Benchmark](D:/Project/TableManager-refactor/benchmarks/corpus_benchmark_2026-11-18.json)
- [Rust 实测](D:/Project/TableManager-refactor/benchmarks/rust_core_benchmark_2026-11-18.json)
