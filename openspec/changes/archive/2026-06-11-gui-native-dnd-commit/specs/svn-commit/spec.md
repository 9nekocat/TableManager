## ADDED Requirements

### Requirement: SVN commit panel
系统 SHALL 提供图形化 SVN 提交面板,显示已修改/新增/删除的文件列表,支持勾选和自定义 commit message。

#### Scenario: 打开提交面板
- **WHEN** 用户点击 toolbar 的 "提交" 按钮
- **THEN** 底部面板展开,显示所有 SVN 变更文件 (M/A/D),每个文件带 checkbox 默认选中,并显示文件路径和变更类型标签

#### Scenario: 输入 commit message 并提交
- **WHEN** 用户勾选文件,输入 commit message,点击 "一键提交"
- **THEN** 系统调用 `svn commit -m "message" file1 file2`,成功后显示 "提交成功 rXXXX",面板刷新

#### Scenario: 未选文件提交
- **WHEN** 用户未勾选任何文件并点击提交
- **THEN** 提交按钮禁用状态,不可点击

#### Scenario: 提交失败
- **WHEN** SVN commit 失败 (网络/锁定/冲突)
- **THEN** 面板显示错误信息,文件选择保持,用户可重试

### Requirement: Commit command injection prevention
SVN commit 的 message 和 file 参数 MUST 通过 `subprocess.run` 的 list 参数传递,禁止使用 `shell=True`。

#### Scenario: 用户输入特殊字符
- **WHEN** 用户在 commit message 中输入 `"; rm -rf /` 或类似内容
- **THEN** 该内容作为 commit message 文本原样传递给 `svn commit -m`,不被 shell 解释

### Requirement: Pre-commit warning
如果本地 SVN 版本落后于远端,系统 SHALL 在提交前显示警告。

#### Scenario: 本地落后远端
- **WHEN** 用户打开提交面板且 `local_rev < remote_rev`
- **THEN** 面板顶部显示黄色警告条 "本地版本 r4520 落后于远端 r4523,建议先更新"

#### Scenario: 冲突文件存在
- **WHEN** 有文件处于 SVN 冲突状态
- **THEN** 冲突文件在列表中标记为红色"C"标签且 checkbox 默认不勾选
