# TableManager 命令与模式参考

## 1. AI 自查：双文件纯文本比对

```powershell
& $tableCore diff $oldPath $newPath --format text
```

只读取 stdout、退出码和 issue 摘要。不要为 AI 自查添加 `--format excel`。

需要结构化处理时使用：

```powershell
& $tableCore diff $oldPath $newPath --format json
```

JSON 仍属于无头模式；只有最终要给人看时才生成 Excel。

## 2. AI 自查：三路冲突检测

```powershell
$probeOutput = Join-Path $env:TEMP "tablemanager-merge-probe.xlsx"
& $tableCore merge $basePath $oursPath $theirsPath `
  --strategy fail --conflict-default theirs --report text --output $probeOutput
$exitCode = $LASTEXITCODE
```

解释：

- `0`：没有真实冲突，允许继续生成用户版报告。
- `2`：存在真实冲突，停止自动合并并等待用户确认。
- 不要添加 `--allow-conflicts` 来做 AI 自查；该选项会生成审查工作簿。
- 如果产生了临时探测文件，自查结束后删除。

## 3. 用户审阅：双文件 Excel 对比

默认排列式报告：

```powershell
& $tableCore diff $oldPath $newPath `
  --format excel --layout arranged --output $reportPath
Start-Process -FilePath $reportPath
```

排列式报告每个原表页签独立，最后生成“比对报告”。若只需旧版紧凑单页样式，必须由用户
明确要求后再使用 `--layout compact`。

SVN 工作副本中的用户对比优先使用：

```powershell
& $svnDiffCmd $basePath $minePath
```

或直接执行 `svn diff`/TortoiseSVN Diff，让 `svn-diff.cmd` 和输出目录配置生效。

## 4. 用户审阅：三路合并预览

SVN/TortoiseSVN 的冲突处理使用：

```text
svn/svn-merge.cmd %merged %theirs %mine %base
```

行为：

- 没有真实冲突：自动写回并标记 `svn resolve`，显示提示框。
- 有真实冲突：生成只读三路预览，按 `[基线]`、`[我的]`、`[仓库]`、`[最终]`
  显示；命令行立即等待用户确认，不轮询 Excel 是否关闭。
- 用户输入 `Y` 后才写回并执行 `svn resolve --accept working`。
- 用户取消时文件继续保持 SVN 冲突状态。

直接使用 CLI 生成用户审查文件：

```powershell
& $tableCore merge $basePath $oursPath $theirsPath `
  --strategy fail --conflict-default theirs `
  --review-output $reviewPath --output $mergedPath

& $tableCore resolve-review $reviewPath $basePath $oursPath $theirsPath `
  --strategy fail --conflict-default theirs --output $mergedPath
```

不要手工编辑 `_MergeMeta`，也不要把 `_MergeReview` 等工作表带入正式业务表。

## 5. 配置

发布包默认配置：`config/svn-wrapper.json`

用户级覆盖：`%LOCALAPPDATA%\CatsInTheShell\TableManager\config\svn-wrapper.json`

常用字段：

```json
{
  "diff_output_dir": "%TEMP%\\TableManager\\SVN",
  "fallback_diff_output_dir": "${SOURCE_DIR}\\.temp",
  "remember_output_dir": true,
  "merge_conflict_default": "theirs"
}
```

`merge_conflict_default`：

- `theirs`：默认采用仓库值。
- `ours`：默认采用本地值。
- `empty`：单元格冲突默认留空；结构冲突仍按仓库侧规则处理。

## 6. 无头环境变量

仅在自动化、CI 或 AI 自查需要调用 SVN 包装器时设置：

```powershell
$env:TM_SVN_HEADLESS = "1"
$env:TM_SVN_NO_OPEN = "1"
$env:TM_SVN_NO_TIP = "1"
```

- `TM_SVN_HEADLESS=1`：不启动外部 UI。
- `TM_SVN_NO_OPEN=1`：不打开生成文件。
- `TM_SVN_NO_TIP=1`：不显示结束提示框。

AI 需要纯文本时，首选直接调用 `table-core.exe diff/merge --format text`，不要依赖
`svn-diff-headless.cmd` 的 Excel 产物。

## 7. 失败处理

- 退出码 `2` 时先保留冲突文件和当前工作表，不得执行提交或覆盖。
- 公式、重复 RowName、重复表头、空行和数据校验问题必须出现在最终报告中。
- 不允许为了消除冲突而静默指定 `--strategy ours` 或 `--strategy theirs`。
- 只有用户明确要求全选某一侧时，才可使用对应策略，并在交付说明中写明。
