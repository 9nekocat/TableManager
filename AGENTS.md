# TableManager 项目更新规则

本文件是 TableManager 的项目级发布流程。处理任何会影响 `table-core` 的改动时，必须
同时考虑 Git 源码仓库和 SVN 导出版仓库，不能把同一份源码或构建产物重复纳入两边。

## 1. 版本库职责

- Git 只管理源码和工程规则：`rust-core/**`、`.gitignore`、`AGENTS.md`。
- SVN 只管理导出版：根目录的 `table-core.exe`、`svn/**`、`config/**`、
  `configure-svn*.bat` 和导出用 `README.md`。
- 禁止执行 `svn add rust-core`。`rust-core` 必须在 SVN 中保持删除或忽略状态。
- 禁止把 `rust-core/target`、`rust-core/dist`、测试输出、日志或压缩包提交到任一仓库。
- 根目录 `README.md`、发布二进制和 SVN 包装器由 SVN 管理。Git 中不得重新跟踪根目录
  `README.md`，也不得使用 `git add -A` 代替选择性暂存。

当前 SVN 忽略项至少应包含：

```text
.git
.gitignore
AGENTS.md
rust-core
target
dist
.temp
*.zip
*.sha256
*.tmp
*.bak
*.log
```

## 2. 修改与验证

所有功能修改先在 `rust-core` 内完成。提交前运行：

```powershell
cargo fmt --manifest-path rust-core/Cargo.toml -- --check
cargo test --locked --manifest-path rust-core/Cargo.toml
cargo clippy --locked --all-targets --manifest-path rust-core/Cargo.toml -- -D warnings
git diff --check
```

涉及 SVN、无头模式或发布脚本时，还必须验证打包后的入口：

```powershell
powershell -ExecutionPolicy Bypass -File rust-core/scripts/package-release.ps1 -Version 0.1.0
```

发布脚本必须通过 `--version`、`install-svn --dry-run`，并确认包内存在：

```text
table-core.exe
configure-svn.bat
config/diff-report.zh-CN.json
svn/svn-diff.cmd
svn/svn-diff-headless.cmd
svn/svn-merge.cmd
svn/svn-merge-headless.cmd
```

## 3. Git 提交流程

只暂存源码和工程规则：

```powershell
git add .gitignore AGENTS.md rust-core
git status --short
git diff --cached --check
git commit -m "feat: describe the completed change"
git push origin <current-branch>
```

根目录导出版即使存在磁盘改动，也不得进入 Git。若发现 Git 仍跟踪根目录导出版，应使用
`git rm --cached <path>` 停止跟踪并同步更新 `.gitignore`，不要删除磁盘上的 SVN 文件。

## 4. 生成并更新 SVN 导出版

Git 源码提交后，从该提交重新生成发布目录：

```powershell
powershell -ExecutionPolicy Bypass -File rust-core/scripts/package-release.ps1 -Version 0.1.0
$stage = "rust-core/dist/table-core-v0.1.0-windows-x64"
```

先更新 SVN，避免基于过期版本覆盖：

```powershell
svn update . --accept postpone
```

然后将发布目录内容同步到 TableManager 根目录，供 SVN 管理：

```powershell
Copy-Item -Path "$stage/*" -Destination "." -Recurse -Force
```

如果 SVN 更新产生了根目录导出版变化，必须在更新后重新执行上述复制，确保 SVN 最终提交
的二进制、脚本、配置和 README 都来自同一个发布版本。

## 5. SVN 提交流程

新增导出版文件时先加入 SVN，但不要加入 `rust-core`：

```powershell
svn add --force --parents table-core.exe configure-svn.bat svn config
svn status
svn diff --summarize
svn commit -m "release: update TableManager export" .
```

如果出现 out-of-date 错误，执行 `svn update . --accept postpone`，重新同步导出版后再提交。
遇到真实冲突时必须保留并检查双方内容，不得直接覆盖远端更新。

## 6. 清理

SVN 提交完成后删除可再生成目录：

```powershell
$workspace = (Resolve-Path '.').Path
foreach ($candidate in @('rust-core/target', 'rust-core/dist')) {
    $path = [IO.Path]::GetFullPath((Join-Path $workspace $candidate))
    if (!$path.StartsWith($workspace + [IO.Path]::DirectorySeparatorChar,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove outside workspace: $path"
    }
    if ([IO.Directory]::Exists($path)) {
        [IO.Directory]::Delete($path, $true)
    }
}
```

同时删除 `.temp`、`*.tmp`、`*.bak`、日志和一次性的 SVN 测试目录。清理后确认：

```powershell
git status --short --branch
svn status
Test-Path rust-core/target
Test-Path rust-core/dist
```

最终状态必须满足：Git HEAD 与远端一致，SVN 无待提交导出版，`rust-core/target` 和
`rust-core/dist` 不存在。

## 7. SVN 一键配置

一键配置脚本的源码位于 `rust-core/configure-svn.bat`，发布时复制为根目录
`configure-svn.bat`。它调用：

```text
table-core.exe install-svn
```

该命令把当前用户 Subversion 的 `diff-cmd` 和 `merge-tool-cmd` 指向当前导出版，并在
修改前保留 `config.table-core.bak`。无头调用使用：

```text
svn/svn-diff-headless.cmd
svn/svn-merge-headless.cmd
```

无头模式不得启动 Excel，也不得回退到 xlCompare。
