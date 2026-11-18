# Merge Permission Lock

合并期间的 SVN 文件锁管理，防止并发修改冲突。

## ADDED Requirements

### Requirement: SVN lock before merge

在执行合并解析（resolve）前，系统 SHALL 对目标文件执行 SVN lock。

#### Scenario: Successful lock before merge
- **WHEN** 用户发起 merge resolve 操作
- **THEN** 系统先调用 `svn lock <file>`，成功后才继续合并

#### Scenario: Lock failure blocks merge
- **WHEN** SVN lock 失败（如文件已被他人锁定）
- **THEN** 系统返回错误信息包含"文件已被锁定"和锁定者信息，不继续合并

### Requirement: SVN unlock after successful merge

合并成功后，系统 SHALL 对目标文件执行 SVN unlock。

#### Scenario: Unlock after successful merge
- **WHEN** merge resolve 成功完成（含 svn resolved）
- **THEN** 系统调用 `svn unlock <file>`，释放文件锁

#### Scenario: Unlock failure warns but merge still succeeds
- **WHEN** 合并成功但 svn unlock 失败
- **THEN** 系统返回成功但附带警告信息，提示用户手动 unlock

### Requirement: Preserve lock on merge failure

合并失败时，系统 SHALL 保留 SVN 文件锁，防止冲突状态被意外覆盖。

#### Scenario: Lock retained on merge failure
- **WHEN** merge resolve 过程中抛出异常
- **THEN** 系统不执行 unlock，保留文件锁
- **THEN** 错误信息中包含文件路径和手动 `svn unlock` 的命令提示

### Requirement: SvnClient lock/unlock methods

`SvnClient` SHALL 提供 `lock(filepath)` 和 `unlock(filepath)` 方法。

#### Scenario: SvnClient.lock returns success
- **WHEN** 调用 `svn.lock("Effect.xlsx")` 且文件未被锁定
- **THEN** 返回 `{"ok": True, "token": "<lock-token>"}`

#### Scenario: SvnClient.unlock releases lock
- **WHEN** 调用 `svn.unlock("Effect.xlsx")` 且当前用户持有锁
- **THEN** 返回 `{"ok": True}`

#### Scenario: Lock by another user detected
- **WHEN** 调用 `svn.lock("Effect.xlsx")` 但文件已被他人锁定
- **THEN** 返回 `{"ok": False, "locked_by": "<username>", "error": "<message>"}`
