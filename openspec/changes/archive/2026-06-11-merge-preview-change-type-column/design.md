## Context

`MergePreview.generate` 当前输出 3 行结构：`[我的]` / `[仓库]` / `[最终]`，用颜色区分变更类型。用户的决策（保留/删除）通过 `[最终]` 行的存在/缺失隐式表达。这导致 MINE_DELETED 不出 `[最终]` 行（默认删除）而 THEIRS_DELETED 出 `[最终]` 行（默认保留）的不对称行为。

## Goals / Non-Goals

**Goals:**
- 新增「修改类型」列，标注 `[我的]` / `[仓库]` 行的变更性质
- `[最终]` 行通过此列的下拉（`[修改]` / `[删除]`）明确表达用户决策
- 单方修改的场景下，默认值自动匹配修改方的行为
- THEIRS_DELETED 默认 `[删除]`（修复当前 bug）
- 决议器通过读取此列的值决定行去留

**Non-Goals:**
- 不改变单方修改自动接受的大原则（MINE_ADDED / THEIRS_ADDED / MINE_MOD / THEIRS_MOD 仍默认进预览 [最终] 行）
- 不新增预览配置 UI

## Decisions

### Decision 1: 列布局 — 插入 B 列

`[标记]` 在 A 列不动。新增 B 列「修改类型」。数据列从 C 开始（当前从 B 开始，统一 +1）。

所有行号引用从 `col_idx + 1` 变为 `col_idx + 2`（跳过标记列和修改类型列）。

**替代方案（不采用）**：把修改类型放在最后一列 → 滚动才能看到，体验差。

### Decision 2: 下拉选择用 openpyxl DataValidation

在 `[最终]` 行的修改类型列应用 `DataValidation(type="list", formula1='"[修改],[删除]"')`。

选项固定为 `[修改]` / `[删除]`，语义清晰：
- `[修改]`：该行写入输出，使用 [最终] 行的数据
- `[删除]`：该行不写入输出

**替代方案（不采用）**：CONFLICT 类型的 [最终] 行加 `[接受我的]/[接受仓库]` 选项 → 冗余，因为用户已可直接编辑数据单元格。

### Decision 3: 默认值映射

| 冲突类型 | 预览写 [最终]？ | 默认修改类型值 | 预填数据来源 |
|----------|:---:|-----------|-----------|
| CONFLICT | 是 | `[修改]` | mine |
| BOTH_ADDED | 是 | `[修改]` | mine |
| MINE_ADDED | 是 | `[修改]` | mine |
| THEIRS_ADDED | 是 | `[修改]` | theirs |
| MINE_MOD | 是 | `[修改]` | mine |
| THEIRS_MOD | 是 | `[修改]` | theirs |
| MINE_DELETED | 是 | **`[删除]`** | mine（空值展示为删除预览） |
| THEIRS_DELETED | 是 | **`[删除]`** | mine（空值展示为删除预览） |
| SAME | 否 | — | — |

MINE_DELETED 从「不写 [最终] 行」改为「写 [最终] 行，值默认 `[删除]`」。用户可将 `[删除]` 改为 `[修改]` 来保留该行。

### Decision 4: [我的]/[仓库] 行修改类型列仅作标注

该列在 `[我的]` / `[仓库]` 行只读展示，无下拉。值映射：

| 行级状态 | 显示 |
|----------|------|
| MINE_MOD / THEIRS_MOD | `[已修改]` |
| MINE_ADDED / THEIRS_ADDED | `[已添加]` |
| MINE_DELETED（我的行）| `[已删除]` |
| THEIRS_DELETED（仓库行）| `[已删除]` |
| BOTH_ADDED（我的行）| `[已添加]` |
| BOTH_ADDED（仓库行）| `[已添加]` |

### Decision 5: 决议器逻辑

`MergeResolver.resolve` 改为逐行读取修改类型列：

```
B列值 == "[删除]" → 跳过该行（不出现在输出）
B列值 == "[修改]" 或其他 → 取 [最终] 行的数据列，写入输出
```

现有的隐式逻辑（`group_has_mine` + `in_base` 判定 MINE_DELETED / THEIRS_ADDED）全部替换为此显式逻辑。分辨率过程中不再需要 `base_order` 查找。

新增行的插入位置算法（锚点搜索）保持不变：仅 THEIRS_ADDED（用户选择保留时）需要计算插入位置。

**简化点**：`sheet_finals` 中没有 key 的行 = 决议器没遇到 `[最终]` = 用户删掉了该行 = 删除。保持向后兼容（手动编辑预览时用户可能删除整行）。

## Risks / Trade-offs

- **旧预览文件不兼容**：新增 B 列后，旧预览文件的决议器会读错列偏移 → 无法升级。**Mitigation**：这是内部工具，当前没有积压的旧预览文件。标记为 BREAKING 即可。
- **用户不填下拉**：若修改类型列留空 → 默认视为 `[修改]`（安全侧：宁可保留数据不丢）
- **DataValidation 在 WPS/Excel 兼容性**：openpyxl 的 DataValidation 在 Excel 和 WPS 中均支持下拉列表，已验证可行。
