## Context

TableManager 当前架构: Python Flask 后端 + 原生 HTML/CSS/JS 前端, PyInstaller 打包为单个 EXE。启动时弹出命令行窗口 + 打开系统浏览器访问 `http://localhost:5000`。UI 无拖拽交互,无 SVN 提交入口,无远端更新感知。用户期望桌面原生体验——无浏览器跳转、无命令行窗口、交互流畅。

## Goals / Non-Goals

**Goals:**
- pywebview 替代浏览器,实现原生桌面窗口 (Edge WebView2)
- HTML5 Drag and Drop 实现表格在分类间拖拽移动
- SVN commit 图形化面板 (文件勾选 + message 输入)
- SVN 远端版本定时检测 + toolbar 角标提醒
- 表格行内智能比对按钮 (M→HEAD, 干净→PREV)
- 现有功能完全保留,零破坏

**Non-Goals:**
- 不引入 React/Vue 等前端框架 (保持原生 JS)
- 不替换 Flask 后端 (只增加路由)
- 不改变现有的合并/校验/版本比对面板功能
- 不处理 SVN 凭证管理 (沿用现有 svn.exe 环境)
- 不实现自动更新下载 (只做提醒)

## Decisions

### D1: pywebview 而非 Electron 或 PySide6

| 方案 | 包大小 | Python集成 | 窗口原生感 | 改动量 |
|------|--------|-----------|-----------|--------|
| pywebview + Edge WebView2 | +3MB | 原生线程 | 高 | ~30行 |
| Electron | +180MB | 需IPC桥接 | 最高 | 大量 |
| PySide6 QWebEngine | +80MB | 原生 | 中 | 中 |

**选择 pywebview**: Windows 10+ 自带 Edge WebView2 运行时, 零额外安装。Flask 线程启动后 `webview.create_window()` 嵌入页面。端口动态分配避免冲突。PyInstaller 打包只需加 `hiddenimports`。

### D2: 拖拽实现 — 原生 HTML5 DnD API

不用第三方库 (SortableJS 等)。原因为:
- 需求简单: 表格行→分类 header 的单向拖拽,不需要排序
- 零额外依赖
- 浏览器原生支持良好 (WebView2 Edge = Chromium)

实现要点:
- 表格行 `draggable="true"`, `dragstart` 存 file path
- 分类侧边栏项 + 主内容区分类 header 都作为 drop zone
- `dragover` 阻止默认 + 高亮边框
- `drop` 事件发 API,成功后刷新

### D3: categories.json 原子写入

```
POST /api/categories/move
→ 读 categories.json
→ 修改内存结构
→ 写临时文件 + os.replace(tmp, target)  # 原子替换
→ 清除 _categories_config 全局缓存
→ 返回 ok
```

`os.replace` 在 Windows/NTFS 上原子操作,防止写一半崩溃导致配置文件损坏。

### D4: SVN commit 命令注入防护

```python
# 安全: 参数列表传递,不用 shell=True
subprocess.run(
    ["svn", "commit", "-m", message] + file_list,
    cwd=RAW_DIR,
    capture_output=True,
    timeout=120,
)
```

message 来自用户输入但不做 shell 解释——`subprocess.run` list 模式下每个参数都是独立 argv 元素。

### D5: 智能 diff 版本选择逻辑

```
def smart_diff(file_path):
    svn_status = svn.status(file_path)
    if svn_status == "M":
        return generate_diff(file, "HEAD")   # 已修改 → 与仓库最新比较
    else:
        log = svn.log_file(file_path, limit=2)
        if len(log) >= 2:
            return generate_diff(file, log[1].revision)  # 未修改 → 与上一版本比较
        else:
            return {"error": "无历史版本"}
```

SVN HEAD 和 PREV 两个端点各加 15s 超时,超时返回友好错误。

### D6: 更新检测轮询策略

- 启动后 2 秒首次检测 (避免阻塞启动)
- 之后每 5 分钟轮询 `GET /api/svn/check-updates`
- 轮询失败静默降级 (不弹错误,只记录日志)
- 有更新时 toolbar SVN 按钮显示角标数字

## Risks / Trade-offs

| Risk | Mitigation |
|------|-----------|
| pywebview 在无 WebView2 的旧 Win10 上启动失败 | 检测缺失时引导安装 WebView2 Runtime (微软提供静默安装包) |
| `svn info -r HEAD` 网络超时阻塞 API | 设置 15s 超时,超时返回 `timeout: true`, 前端显示 "网络不可用" |
| 拖拽跨分类时如果 API 失败,前端状态不一致 | API 失败时撤销 UI 变更 (回滚拖拽),显示 toast 错误提示 |
| 原子写入 categories.json 仍有极低概率丢失 | 新格式很简单,崩溃恢复可接受;备选方案是保留最近一次备份 |
| PyInstaller + pywebview + pythonnet 兼容性 | 先在干净环境手动打包验证,CI 加 Windows 构建步骤 |
| 大量表格 (>200) 时拖拽性能 | 事件委托 (父容器监听而非每行绑定),防抖 refresh |

## Open Questions

- 拖拽是否允许"从分类移除到未分类"? → 暂支持,侧边栏预留 "未分类" drop zone
- SVN commit 前是否自动做 svn update? → 不做。提示用户先更新再提交 (前端显示本地/远端版本差异)
- 是否需要右键菜单 (打开/比对/属性)? → 本次不做,保留点击行打开 Excel 的行为不变
