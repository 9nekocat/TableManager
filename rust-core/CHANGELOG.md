# Changelog

## Unreleased

- 合并预览改为按原表页签组织 `[基线]`、`[我的]`、`[仓库]`、`[最终]` 行。
- 新增冲突回写链路：`merge --review-output` 生成预览，`resolve-review` 应用
  `[最终]` 行并输出干净的业务工作簿。
- `svn-merge` 在无真实冲突时自动合并并执行 `svn resolve --accept working`。
- 有真实冲突时等待用户关闭 Excel 并在命令窗口确认后写回、清理预览并标记 resolved。
- 非无头模式在合并完成或取消后显示 15 秒提示框，避免命令窗口自动关闭导致用户看不到结果。
- 冲突默认值改为仓库侧 `theirs`，可通过 `merge_conflict_default` 配置为
  `ours` 或 `empty`。

## 0.1.0 - 2026-11-18

- 首个独立的 Rust Excel 比对与三路合并应用。
- 支持文本、JSON 和 Excel 差异报告。
- Excel 差异报告使用单页布局，摘要置顶，变更内容保留原表列头与单元格位置。
- 差异报告辅助文案支持中文本地化 JSON 配置覆盖。
- 新增 `install-svn` 自注册命令，将 Subversion xlsx 族比对与合并指向当前发布包。
- 新增 SVN 无头包装脚本，支持不启动 Excel 的自动化比对与合并。
- 一键 SVN 配置脚本同时更新 Subversion 命令行和 TortoiseSVN 的 xlsx 族 diff/merge 配置，并提供管理员模式入口。
- 新增 Excel 排列模式：每个原表页签独立，末尾生成问题汇总报告，并压缩同类空行问题。
- SVN diff 默认将报告写入工作文件旁的 `.temp`，避免 `%TEMP%` 写入受限时命令窗口直接退出。
- SVN diff 输出目录支持首选/回退配置，并会记忆成功回退路径。
- 支持单元格、行、列和工作表的增删改。
- 支持简单公式保留与行内引用重绑定。
- 提供冲突审查工作簿和明确的退出码。
- 提供 SVN 命令行与 TortoiseSVN 包装脚本。
