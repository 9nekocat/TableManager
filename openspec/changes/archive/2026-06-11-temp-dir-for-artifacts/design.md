## Context

当前 `_diff_r*.xlsx` 和 `_merge.xlsx` 直接生成在 `RAW_DIR` 下，与源数据文件混在一起。需统一移到 `RAW_DIR/.temp/`。

## Goals / Non-Goals

**Goals:**
- 所有 diff/merge 临时文件输出到 `RAW_DIR/.temp/`
- 使用 `RAW_DIR` 相对路径构建，保持可移植性

**Non-Goals:**
- 不改变 resolve 后的最终输出路径（仍在 `RAW_DIR` 下）
- 不增加自动清理逻辑（后续可加）

## Decisions

**统一常量**: 在 `manager_core.py` 定义 `TEMP_DIR = RAW_DIR / ".temp"`，各模块引用。

**相对路径结构**: 镜像源文件在 `RAW_DIR` 下的结构：
```
RAW_DIR/.temp/Effect_diff_r4583.xlsx
RAW_DIR/.temp/Effect_merge.xlsx
```

**api_merge_resolve 原文件推断**: 当前用 `str(preview_path).replace("_merge.xlsx", ".xlsx")` 字符串替换。preview 移到 `.temp/` 后此逻辑失效，改为基于 `RAW_DIR` 构造：从 preview 文件名去掉 `_merge` 后缀，拼接 `RAW_DIR / base_rel`。

## Risks / Trade-offs

- `.temp/` 目录需加入 `.svnignore`，否则 SVN 会看到临时文件
