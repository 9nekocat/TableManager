## ADDED Requirements

### Requirement: Native desktop window via pywebview
系统 SHALL 使用 pywebview 将 Web 前端嵌入原生桌面窗口,使用 Edge WebView2 引擎渲染。启动时 MUST NOT 弹出命令行窗口或外部浏览器。

#### Scenario: 正常启动
- **WHEN** 用户双击 TableManager.exe
- **THEN** 系统启动 Flask 服务器于随机可用端口,弹出原生窗口 (默认 1280x860),加载首页,无浏览器地址栏,无命令行黑窗

#### Scenario: WebView2 不可用
- **WHEN** 系统未安装 Edge WebView2 运行时
- **THEN** 系统显示错误对话框提示安装 WebView2 Runtime,并退出

#### Scenario: 端口冲突处理
- **WHEN** 默认端口被占用
- **THEN** 系统自动选择下一个可用端口,并在窗口标题栏显示实际端口号

### Requirement: Window lifecycle
原生窗口 SHALL 在关闭时自动停止 Flask 服务器线程并退出进程。

#### Scenario: 用户关闭窗口
- **WHEN** 用户点击窗口关闭按钮
- **THEN** Flask 服务器线程终止,进程退出,不残留后台进程

### Requirement: Backward compatible CLI mode
系统 MUST 保留纯 Web 服务器模式,以便开发调试或无法使用 WebView2 的环境。

#### Scenario: 命令行启动 Web 模式
- **WHEN** 用户执行 `python table_server.py --web`
- **THEN** 系统以传统 Flask HTTP 模式启动,打印 URL,不创建原生窗口
