# Batch Excel Recalc

一次 Excel COM 会话中重算多个 xlsx 文件，避免逐个文件启动独立 Excel 进程。

## ADDED Requirements

### Requirement: Batch recalc multiple files in one session

`recalc_batch(filepaths)` SHALL 在一次 Excel COM 会话中打开所有文件、解析链接工作簿、执行全量重算、保存并关闭。

#### Scenario: Two files recalc in one session
- **WHEN** 调用 `recalc_batch(["Effect_previous.xlsx", "Effect_latest.xlsx"])`
- **THEN** 两个文件在同一个 Excel 实例中被打开、链接、重算、保存
- **THEN** 重算后的文件包含最新的公式缓存值

#### Scenario: Linked workbooks resolved once
- **WHEN** 两个文件共享相同的链接工作簿
- **THEN** 链接工作簿只打开一次（去重），不重复打开

#### Scenario: Originals not modified
- **WHEN** `recalc_batch` 被调用时传入的是临时副本
- **THEN** 只有临时副本被修改，原始文件保持不变

### Requirement: Graceful fallback on COM failure

COM 不可用时 SHALL 静默跳过重算，返回原文件路径列表。

#### Scenario: win32com not available
- **WHEN** `recalc_batch` 执行时 `win32com` 未安装
- **THEN** 函数返回原始 filepaths 列表，不抛出异常
