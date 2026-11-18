## 1. pywebview 集成

- [x] 1.1 添加 `pywebview` 和 `pythonnet` 到 requirements.txt
- [x] 1.2 在 `table_server.py` 新增加载 webview 入口: Flask 线程 + `webview.create_window()` + `webview.start()`
- [x] 1.3 动态端口分配: 端口被占用时自动+1 重试
- [x] 1.4 窗口关闭时 signal 事件停止 Flask 线程 → 进程退出
- [x] 1.5 保留 `--web` CLI 参数: 纯 Flask HTTP 模式供开发调试
- [x] 1.6 更新 `TableManager.spec`: `console=False`, `hiddenimports` 加 `webview`, `webview.platforms.winforms`, `clr`

## 2. 新增后端 API

- [x] 2.1 `POST /api/categories/move`: 原子读→改→写 `categories.json`, 清除缓存, 返回 ok
- [x] 2.2 `POST /api/svn/commit`: 校验参数 (message 非空, files 列表非空), `subprocess.run` list 模式调用 svn, 返回 revision
- [x] 2.3 `GET /api/svn/check-updates`: `svn info` 取本地 rev, `svn info -r HEAD` 取远端 rev, 返回 diff
- [x] 2.4 `GET /api/diff/smart/<path>`: svn status 判断 M→调用 diff vs HEAD, 否则取 prev rev→调用 diff vs PREV
- [x] 2.5 所有新 API 加 try/except + 超时保护 (svn 命令 15-30s 超时)

## 3. SVN 客户端扩展

- [x] 3.1 `svn_client.py` 新增 `commit(files, message)` 方法
- [x] 3.2 `svn_client.py` 新增 `check_updates()` 返回 local_rev, remote_rev, has_update
- [x] 3.3 `svn_client.py` 新增 `get_file_prev_rev(file_path)` 取倒数第二次修改的版本号
- [x] 3.4 `diff_engine.py` 新增 `smart_diff(file_path)` 便捷函数

## 4. 前端: 拖拽移动分类

- [x] 4.1 表格行加 `draggable="true"`, `dragstart` 事件存储 `{file, catKey}`
- [x] 4.2 侧边栏分类项 + 主内容区分类 header 注册 `dragover`/`dragleave`/`drop` 事件
- [x] 4.3 drop zone 高亮样式 (`.cat-drop-target`): 2px accent 虚线边框 + 浅色背景
- [x] 4.4 拖拽幽灵图 + 半透明效果
- [x] 4.5 drop 事件处理: 同分类跳过, 不同分类发 `POST /api/categories/move`, 成功刷新, 失败 toast 提示并回滚

## 5. 前端: SVN 提交面板

- [x] 5.1 toolbar 新增 "提交" 按钮 (有变更时高亮)
- [x] 5.2 底部提交面板: 变更文件列表 (checkbox + 文件名 + 类型标签 M/A/D/C)
- [x] 5.3 commit message 输入框 (textarea, 单行输入)
- [x] 5.4 "一键提交" 按钮: 未选文件时 disabled, 点击后 loading → 成功/失败反馈
- [x] 5.5 提交前 pre-commit 检查: 本地落后远端时黄色警告条
- [x] 5.6 提交成功后自动刷新 SVN 状态 + 关闭面板

## 6. 前端: SVN 更新提醒

- [x] 6.1 `loadSvnCheckUpdates()`: 调用 `/api/svn/check-updates`, 超时静默处理
- [x] 6.2 toolbar SVN 按钮右上角角标 (圆点/数字, CSS `position: absolute` 红色圆 badge)
- [x] 6.3 启动时 2 秒后首次检测, 之后 `setInterval(300000)` 每 5 分钟
- [x] 6.4 有更新时 hover tooltip "远端有新版本 rXXXX, 落后 N 个版本"

## 7. 前端: 行内智能比对

- [x] 7.1 表格行 `col-actions` 区新增 [比对] 按钮 (紧凑样式,不干扰[打开]按钮)
- [x] 7.2 `quickDiff(file)`: 调用 `GET /api/diff/smart/<file>`, loading 态处理
- [x] 7.3 结果复用现有 `renderDiffResults()` 展示在 `diff-panel`
- [x] 7.4 右上角 "版本比对" 按钮保留不变 (手动高级模式)

## 8. 前端微优化

- [x] 8.1 拖拽过渡动画: CSS `transition` 分类边框颜色 150ms
- [x] 8.2 toast 通知条: 固定底部居中, 成功(绿色)/错误(红色), 3秒自动消失
- [x] 8.3 空状态占位: 分类下无表格时显示 "此分类无表格"
- [x] 8.4 行 hover 过渡: 背景色 `transition 120ms`
- [x] 8.5 表格行骨架屏: 首次加载时显示占位条而非仅 loading spinner

## 9. 打包与验证

- [x] 9.1 本地测试 pywebview 窗口模式: `python table_server.py`
- [x] 9.2 PyInstaller 打包: `pyinstaller TableManager.spec`
- [x] 9.3 验证 EXE 启动: 无 console 窗口, 原生窗口正常, WebView2 正常加载
- [x] 9.4 测试拖拽流程: 拖→放→配置文件更新→UI 刷新
- [x] 9.5 测试 SVN commit 流程: panel→勾选→message→commit→成功
- [x] 9.6 测试智能比对: 已修改文件 + 未修改文件 + 非 SVN 文件 三种情况
- [x] 9.7 测试 `--web` 模式: 传统 Flask HTTP 模式仍可启动
