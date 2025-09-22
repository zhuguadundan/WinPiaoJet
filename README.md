# 发票 A5 到 A4 上半张打印工具

本工具面向财务/行政等场景，支持将 A5（含横向/纵向/准 A5）PDF 发票批量打印到 A4 纸“上半张”固定区域，并提供常用 PDF 工具（图片→PDF、PDF→图片、合并/拆分、旋转、水印/页码）。

## 主要能力
- 打印：A5→A4 上半张，1:1 优先，硬件边距不足时 90%–100% 自动微缩。校准测试页、模板（边距/安全距/微缩）
- 预览：显示可打印区域、上半区域、安全距、毫米标尺与缩放/尺寸
- 批量：文件/文件夹/拖拽导入、进度与失败重试、CSV 导出
- 工具：
  - 图片→PDF（A4/A5、边距、等比）
  - PDF→图片（PNG/JPG/TIFF、150/300/600DPI、JPEG 质量）
  - 合并/拆分（按顺序合并、页区间拆分）
  - 旋转（90/180/270）、文本水印（中心/斜向、透明度）、页码（格式/位置）

## 运行与发布
- 运行：使用 .NET 8（Windows 10/11）
  - `dotnet build` 后运行 `src/Pdftools.Desktop/bin/Debug/net8.0-windows/Pdftools.Desktop.exe`
- 发布自包含：
  - 运行 `scripts/publish.ps1`（输出在 `dist/win-x64`）

## GitHub Release（自动化）
- 已配置 GitHub Actions：推送符合 `v*` 的标签将自动构建并创建 Release，附带以下产物：
  - `dist/win-x64/winpiaojet.exe`
  - `dist/winpiaojet-win-x64.zip`
  - Release 说明会自动从 `CHANGELOG.md` 中抽取对应版本的节选（若无则回退为提交摘要）
- 发布流程：
  - 更新版本号（可选）：编辑 `Directory.Build.props` 中的 `VersionPrefix`
  - 打标签并推送：
    - `git tag v1.0.1`
    - `git push origin v1.0.1`
  - 稍等片刻，Release 将在 GitHub 上生成

## 右键菜单集成（Windows）
- 发布包已内置脚本：解压后在 `tools/` 目录下可找到：
  - `install-context-menu.ps1`
  - `uninstall-context-menu.ps1`
- 安装（在发布包根目录执行，ExePath 指向同级 exe）：
  - `powershell -ExecutionPolicy Bypass -File tools\install-context-menu.ps1 -ExePath .\winpiaojet.exe`
- 卸载：
  - `powershell -ExecutionPolicy Bypass -File tools\uninstall-context-menu.ps1`
- 细节说明见：`docs/右键菜单集成.md`（注册表位置、多选、静默打印说明）

## 打印校准与建议
- 首次使用请打印“测试页”进行校准；根据实际打印结果微调“边距/安全距”并保存为模板
- 驱动设置建议：关闭“适合页面/缩放”自动缩放，纸张 A4，单面

<!-- 压缩功能已移除 -->

## 日志与问题排查
- 日志目录：`%LocalAppData%\pdftools\logs`，按天滚动
- 崩溃/异常会写入日志；可在“导出 CSV”获取批量任务状态
- 详见 `docs/问题排查.md`

## 许可证
- 仅使用 BSD/MIT/Apache-2.0 等宽松许可依赖（PDFium、PDFsharp、SkiaSharp 等）。
