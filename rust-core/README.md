# table-core-rs

无界面的 Rust 核心，供 CLI、Python GUI 或其他程序复用。

职责：

- 读取 `.xlsx` 的 `#` 数据页。
- 把表冻结成纯文本快照，公式只取 Excel 缓存值。
- 检查表头、RowName、Excel 错误值和公式风险。
- 三路合并，支持 `fail`、`ours`、`theirs` 冲突策略。
- 合并时保留简单公式，并按输出行列重绑定行内引用。
- 输出类 Git 文本、JSON 或差异 Excel。
- 输出合并 Excel，附带 Review/Conflicts/Changes/Issues 审查页。

不包含 GUI、SVN、Excel COM 和原地事务写回。复杂外链、动态数组元数据和复杂格式保留仍复用现有 Python 链路。

```powershell
cargo build --release
.\target\release\table-core.exe validate A:\DataTable\Raw\Effect.xlsx --json
.\target\release\table-core.exe diff old.xlsx new.xlsx --format text
.\target\release\table-core.exe diff old.xlsx new.xlsx --format excel --output diff.xlsx
.\target\release\table-core.exe merge base.xlsx ours.xlsx theirs.xlsx --strategy fail --output merged.xlsx
```
