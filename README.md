# DataTable Manager

UE DataTable 配置表管理工具。提供表格浏览、校验、SVN 更新、Excel 冲突比对与合并功能。

## 重构中的核心

`src/core/` 提供不依赖 GUI、SVN 和 Excel COM 的纯数据核心，支持校验、CRUD、两文件差异和三路合并。差异会先把工作簿转成纯文本快照，公式直接使用 Excel 缓存值。设计与审查结论见 [docs/core-refactor-review.md](docs/core-refactor-review.md)，真实表性能见 [docs/real-table-benchmark-rust-estimate.md](docs/real-table-benchmark-rust-estimate.md)。

高性能只读路径已经迁移到 Rust：`rust-core/` 提供无界面的库和 `table-core` CLI，覆盖校验、文本快照、差异及 JSON/文本/Excel 输出。

开发入口：

```powershell
python src/core_cli.py validate <file.xlsx> --json --clean
python src/core_cli.py diff <old.xlsx> <new.xlsx> --format text
python src/core_cli.py merge <base.xlsx> <ours.xlsx> <theirs.xlsx> --output <merged.xlsx>

cargo build --release --manifest-path rust-core/Cargo.toml
rust-core\target\release\table-core.exe diff <old.xlsx> <new.xlsx> --format text
```

## 环境准备

### 1. SlikSVN（SVN 命令行工具）

TableManager 的 SVN 更新及冲突检测依赖 `svn.exe`。TortoiseSVN 默认不带命令行工具，需单独安装 SlikSVN：

```
env\Slik-Subversion-1.14.5-x64.msi
```

启动 TableManager 时会自动检查 `svn.exe`。如果未安装，程序会优先运行 `env` 目录下的 SlikSVN 安装包；安装后会重新检测 `svn.exe`。

也可以手动双击安装，全部默认选项即可。安装后 `svn.exe` 自动加入 PATH。

> 如果已通过其他方式安装了 SVN 命令行工具（如 VisualSVN、Apache Subversion），可跳过此步。

### 2. Excel

需安装 Microsoft Excel 或 WPS（用于打开合并预览文件）。

## 快速开始

双击 `TableManager.exe`，自动打开浏览器访问 Web 管理界面。

### Web 界面功能

| 功能 | 说明 |
|------|------|
| 表格列表 | 按类别浏览 DataTable/Raw 下所有表格 |
| 搜索 | 按文件名或 UE 表名搜索 |
| 校验 | 扫描所有表格中的 Excel 错误值（#N/A, #REF! 等） |
| SVN 状态 | 查看 DataTable/Raw 的 SVN 变更 |
| SVN 更新 | 拉取最新版本 |
| 冲突合并 | 检测冲突文件，生成合并预览，应用合并结果 |
| 妙妙工具箱 | ① Effect #N/A 修复重算（公式 OR 保护 + COM 全算，消除空单元格导致的 #N/A）；② 查 #N/A（对任意 Raw 表扫描公式错误缓存并列出明细） |

## 合并工作流

当 `svn update` 后出现 Excel 文件冲突时：

```
1. 点击 "SVN 状态" → "更新"
2. 更新完成后自动加载冲突列表
3. 点击冲突文件的 "合并预览"
4. 在 Excel 中查看并编辑 [最终] 行的值
5. 保存关闭 Excel
6. 点击 "应用合并"
```

### 合并预览文件说明

```
[标记]  |  行名  |  伤害  |  速度
(type)  | string |  int   | float
--------|--------|--------|-------
[我的]  | IceBolt|   60   |  2.0    ← 本地版本
[仓库]  | IceBolt|   70   |  2.0    ← 服务器版本
[最终]  | IceBolt|   60   |  2.0    ← 用户决策（预填 mine 值）
```

- **[我的]**: 黄色背景 — 本地修改的版本
- **[仓库]**: 黄色背景 — 服务器版本
- **[最终]**: 绿色背景 — 用户直接编辑此行，决定最终值
- **冲突单元格**: 红色高亮
- **公式**: 显示为 `'=FORMULA → VALUE` 格式（Excel 文本模式）

### 合并规则

| 场景 | [我的] | [仓库] | [最终] 预填 |
|------|--------|--------|------------|
| 双方修改不同 | 显示 | 显示 | mine 值 |
| 仅本地修改 | 显示 | - | mine 值 |
| 仅服务器修改 | - | 显示 | theirs 值 |
| 本地新增 | 显示 | - | mine 值 |
| 服务器新增 | - | 显示 | theirs 值 |
| 无冲突 | 不显示 | 不显示 | 自动保留 |

## 命令行用法

```bash
# 表格管理
TableManager.exe list               # 列出所有表格
TableManager.exe list -c Combat     # 按类别筛选
TableManager.exe open Effect        # 在 Excel 中打开
TableManager.exe info CardTable     # 查看详情
TableManager.exe check              # 校验全部
TableManager.exe check -f Effect    # 校验单个文件

# SVN 操作
TableManager.exe svn status         # 查看 SVN 状态
TableManager.exe svn update         # 更新 DataTable/Raw
TableManager.exe svn log -n 10      # 查看最近日志

# 合并操作
TableManager.exe merge status       # 列出冲突文件
TableManager.exe merge preview Effect   # 生成合并预览
TableManager.exe merge resolve Effect   # 应用合并结果
TableManager.exe merge abort Effect     # 放弃合并
```

## 目录结构

```
TableManager/
├── TableManager.exe        # 主程序（PyInstaller 打包产物，位于 dist/）
├── env/
│   └── Slik-Subversion-1.14.5-x64.msi   # SVN 命令行工具安装包
├── gui/                     # React 前端（src/ 源码，dist/ 构建产物）
├── categories.json         # 表格分类配置
├── art/                    # 图标与字体资源
└── README.md
```

## 配置

### categories.json

编辑此文件可自定义表格分类：

```json
{
  "categories": {
    "Combat": {
      "display_name": "战斗",
      "description": "Effects, targets, buffs",
      "files": [
        "Effect.xlsx",
        "Target.xlsx"
      ]
    }
  }
}
```

### 项目根目录检测

TableManager 自动向上查找包含 `DataTable/Raw` 或 `.uproject` 的目录作为项目根目录。也可设置环境变量 `CATSINTHESHELL_ROOT` 指定。

## 常见问题

**Q: 提示 "SVN not found"？**
A: 安装 `env\Slik-Subversion-1.14.5-x64.msi`。

**Q: 合并预览中公式显示为 `#NAME?`？**
A: 公式以 `'=` 开头（Excel 文本模式），显示为文本而非计算结果。合并后公式会恢复为真实 Excel 公式。

**Q: 合并后非冲突行丢失？**
A: 不会。合并以 mine 文件为基底，仅替换冲突行的值，所有非冲突行和 sheet 完整保留。

**Q: 打开文件失败？**
A: 确保已安装 Excel 或 WPS。
