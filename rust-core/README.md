# table-core-rs

无界面的只读 Rust 核心，供 CLI、Python GUI 或其他程序复用。

职责：

- 读取 `.xlsx` 的 `#` 数据页。
- 把表冻结成纯文本快照，公式只取 Excel 缓存值。
- 检查表头、RowName、Excel 错误值和公式风险。
- 输出类 Git 文本、JSON 或差异 Excel。

不包含 GUI、SVN、Excel COM、事务写回、复杂外链/动态数组保留。这些继续复用现有 Python 链路。

```powershell
cargo build --release
.\target\release\table-core.exe validate A:\DataTable\Raw\Effect.xlsx --json
.\target\release\table-core.exe diff old.xlsx new.xlsx --format text
.\target\release\table-core.exe diff old.xlsx new.xlsx --format excel --output diff.xlsx
```
