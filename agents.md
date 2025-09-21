# AGENTS.md（项目工作指引 / 项目结构与技术栈）

## 项目目标
- 面向财务/行政，将 A5 发票精确打印到 A4 上半张，并提供常用 PDF 工具（图片→PDF、PDF→图片、合并/拆分、旋转、水印/页码）。

## 架构与技术栈
- 运行时/语言：.NET 8（Windows，仅 x64）
- 桌面 UI：WPF（`net8.0-windows10.0.19041.0`，启用 `UseWPF`/`UseWindowsForms`）
- PDF 生成：PDFsharp（轻量生成/绘图）
- PDF 渲染：Windows.Data.Pdf（WinRT 渲染到内存流，供预览）
- 日志：Serilog（File sink，按天滚动，目录 `%LocalAppData%/pdftools/logs`）
<!-- 压缩依赖已取消 -->
- 打包发布：PowerShell 脚本 `scripts/publish.ps1`（自包含单文件，输出到 `dist/win-x64` 并 Zip）

## 目录结构
- `src/Pdftools.Core`：通用能力与工具
  - `Utilities/Paths.cs`：应用/日志目录
  - `PdfOps/*`、`Models/*`、`Services/*`：核心功能（按需扩展）
- `src/Pdftools.Desktop`：WPF 客户端
  - `App.xaml / App.xaml.cs`：启动与自检、全局异常
  - `Services/WindowsPdfRenderer.cs`：WinRT PDF 渲染
  <!-- 压缩模块调用器已移除 -->
  - `Windows/*`、`MainWindow.xaml*`：窗口/视图
- `scripts/publish.ps1`：发布到 `dist/win-x64` 并打包 zip
- `dist/win-x64`：发布产物文件夹
- `docs/*`：用户手册与问题排查

## 构建与运行
- 构建：`dotnet build Pdftools.sln -c Debug`
- 调试：VSCode 配置见 `.vscode/launch.json`
- 运行：`src/Pdftools.Desktop/bin/Debug/net8.0-windows10.0.19041.0/Pdftools.Desktop.exe`

## 发布
- 脚本：`powershell -ExecutionPolicy Bypass -File scripts/publish.ps1 -Configuration Release`
- 产物：`dist/win-x64/Pdftools.Desktop.exe` 与 `dist/pdftools-win-x64.zip`
- 注意：如发布目标 EXE 正在运行，会被占用。脚本已增加检测；可用 `-KillRunning` 参数强制结束运行中的同路径实例。

## 日志与排障
- 日志路径：`%LocalAppData%/pdftools/logs/app-*.log`
- 未处理异常：UI/非 UI/任务异常均写入日志，并给出用户提示
<!-- 压缩相关功能已移除 -->

## 关键约定
- 自检不落磁盘，使用内存流生成/渲染，避免文件映射占用冲突
- 代码与注释统一使用中文，最小必要改动、聚焦问题根因
- 服务/工具调用失败时记录日志并优雅降级，不阻塞主流程

## 常见风险与对策
- 发布占用：目标 EXE 被占用导致发布失败 → 脚本启动前检测并提示；可选 `-KillRunning`
- PDF 预览：字体解析引发依赖冲突 → 自检仅画图形；业务侧尽量避免未嵌入字体的复杂文本渲染
<!-- 不再依赖压缩外部工具 -->

## 贡献与风格
- 保持方法短小、职责单一；异常捕获时使用语义化日志信息
- 变更前先阅读相关模块与调用链；跨文件改动须在 PR 描述中标注影响面
- 不额外引入重量级依赖；优先复用现有工具链
