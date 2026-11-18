---
name: effect-check-config
description: 配置 TableManager 的导表检查与引用查找规则。当需要新增/修改跨表引用规则（references.json）、Effect 表 Type 参数语义（effect_checks.json）、或校准 categories.json 时使用。触发词：「引用规则」「导表检查配置」「Effect 参数规则」「脚本配置方法」「新增 Type 规则」。
---

# 导表检查 / 引用查找 规则配置 Skill

本 skill 指导如何维护 TableManager 的两份规则配置：

| 文件 | 作用 |
|---|---|
| `references.json` | 通用跨表引用规则：列级 ID 引用 + 文本内嵌 `{Xxx,ID}` 引用 |
| `effect_checks.json` | Effect.xlsx 专用：按 Type 的 Params 位置参数语义 |

规则语义的一手来源是**小闪在线文档[脚本配置方法]**；`A:\DataTable\Raw\配置表使用方法.docx` 有各表字段说明。改规则前先查这两处，再用数据实证（见下文"验证方法"）。

## references.json

```jsonc
{
  "column_rules": [
    {
      "name": "规则名（报告显示用）",
      "file": "来源表.xlsx",            // 规则作用哪个文件
      "sheets": ["#Main"],              // 可选，限定 sheet
      "column_pattern": "^EffectId$",   // 正则，同时匹配第1行中文名和第2行英文名
      "array_sep": ",",                 // 可选，数组拆分符，默认 ","；有用 "_" 的列
      "targets": [                      // 候选目标列表，任一命中即合法
        { "file": "Effect.xlsx" },
        { "file": "Localization/Tag.xlsx", "sheets": ["#Stage"], "column": "Tag" }
        // sheets 可选（默认全部 # sheet）；column 可选（默认 A 列 RowName，可指定其他列如 Tag）
      ]
    }
  ],
  "text_patterns": [
    {
      "name": "Buff内嵌引用",
      "regex": "\\{Buff[,，]\\s*(\\d+)\\}",   // 捕获组1 = ID，扫描所有表的字符串单元格
      "targets": [{ "file": "Localization/Buff.xlsx" }]
    }
  ]
}
```

注意：
- 很多列的第 2 行英文名是字面量 `"None"`，必须靠第 1 行中文名匹配（如 `^(击中|非击中)效果ID\d+$`）。
- 多态 ID 列（如卡牌效果 ID：9xxxxxx→Effect、8xxxxxx→机师技能）用多候选 targets 解决。
- 位置参数列（一个单元格里不同位置语义不同）**不要**用 column_rules，Effect 表走 effect_checks.json。

## effect_checks.json

Effect.xlsx 的 `Params` 列是位置参数，含义由 `Type` 列决定：

```jsonc
{
  "file": "Effect.xlsx",
  "condition_targets": [{ "file": "Localization/ShopFilterCondition.xlsx" }],  // 条件列全局目标
  "types": {
    "9": {                                   // Type 值（字符串 key）
      "name": "事件监听（事件+目标+子效果）",   // 文档用途
      "level": "error",                      // 可选，悬空严重度，默认 error
      "params": [
        { "pos": 1,   "name": "事件ID", "kind": "ref", "targets": [{ "file": "Event.xlsx" }] },
        { "pos": 2,   "name": "目标ID", "kind": "ref", "targets": [{ "file": "Target.xlsx" }] },
        { "pos": "3+","name": "效果ID", "kind": "ref", "targets": [{ "file": "Effect.xlsx" }] }
        // pos: 数字 或 "N+"（第 N 个及以后全部）
        // kind: "ref" = 做悬空检查；"literal" = 字面值，仅作文档不检查
      ]
    }
  }
}
```

**新增/修正一个 Type 规则的流程**：
1. 查小闪文档[脚本配置方法]确认该 Type 每个参数位置的语义；
2. 数据实证（必须先做，文档可能滞后）：跑项目根目录的一次性统计——按 Type 分组，统计每个参数位置的值命中各候选表 ID 空间的比例（Event/Target/Effectnumber/Projectile/Buff/Card/ShopFilter/Effect 自身）。命中率 ≥95% 且 ID 空间不歧义才可定为 `ref`；
3. **小数值陷阱**：1、2、3 这类字面量会撞 Projectile/Card 等小 ID 空间，命中率虚高——这种位置应标 `literal`；
4. 写入 `types`，未知语义的 Type 保留 `"name": "待确认（见小闪文档[脚本配置方法]）"`、`"params": []`；
5. 证据不足 95% 但确实有引用语义的，加 `"level": "warning"`。

## 验证方法

```bash
# 改完规则后必须跑
python -m pytest tests/test_ref_index.py -v   # 规则解析单测
python table_manager.py check -f Effect       # Effect 表单表检查（看悬空引用数是否合理）
python table_manager.py refs --dangling       # 全量悬空引用清单
python scan_categories.py                     # categories.json 与磁盘是否一致
```

- 引用索引按文件 mtime 缓存，改表后自动重扫；改 json 配置后**重启进程**生效（CLI 每次是新进程，GUI 需重启服务）。
- 悬空引用突然大量增加，通常说明：目标表 ID 空间变了（先跑 `scan_categories.py` 看是否缺文件）、或规则的目标表配错。

## 相关文件

- 扫描/索引实现：`ref_index.py`；检查规则实现：`table_checks.py`
- 导出数据（Processed/DT_*.json）检查：`processed_checks.py` + `processed_checks.json`——位置参数规则按**导出枚举名**（如 `ECardEffectType::Event`）配置，与本 skill 的数字 Type 规则一一对应；`target` 用逻辑表名（`{"table": "Effect"}`），`target.column` 同样可用。注意 `'0'` 是通用"无"占位符，引擎已统一跳过
- 测试：`tests/test_ref_index.py`、`tests/test_table_checks.py`、`tests/test_processed_checks.py`
- categories.json 校准工具：`scan_categories.py`（`--write` 写回）
