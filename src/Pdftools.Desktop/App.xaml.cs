using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using WpfApplication = System.Windows.Application;
using Serilog;
using Pdftools.Core.Utilities;

// 自检所需
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using Pdftools.Desktop.Services;
using PdfSharpDoc = PdfSharp.Pdf.PdfDocument;
using Windows.Storage.Streams;

namespace Pdftools.Desktop
{
    /// <summary>
    /// 应用程序入口（WPF）
    /// </summary>
    public partial class App : WpfApplication
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            Directory.CreateDirectory(Paths.LogsDir);
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(Path.Combine(Paths.LogsDir, "app-.log"), rollingInterval: RollingInterval.Day)
                .CreateLogger();

            // 全局异常捕获
            this.DispatcherUnhandledException += (s, exArgs) =>
            {
                Log.Error(exArgs.Exception, "UI 未处理异常");
                exArgs.Handled = true;
                System.Windows.MessageBox.Show("发生未处理异常，已记录到日志。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (s, exArgs) =>
            {
                if (exArgs.ExceptionObject is Exception ex)
                    Log.Error(ex, "非 UI 未处理异常");
                else
                    Log.Error("非 UI 未处理异常: {ExceptionObject}", exArgs.ExceptionObject);
            };
            TaskScheduler.UnobservedTaskException += (s, exArgs) =>
            {
                Log.Error(exArgs.Exception, "任务未观察异常");
                exArgs.SetObserved();
            };

            Log.Information("应用启动");

            // 解析命令行参数
            bool toPrint = false;
            string? printer = null; string? templateId = null; int dpi = 300;
            string[] pdfs = Array.Empty<string>();
            try
            {
                var args = e.Args ?? Array.Empty<string>();
                toPrint = args.Any(a => string.Equals(a, "--print", StringComparison.OrdinalIgnoreCase) || string.Equals(a, "/print", StringComparison.OrdinalIgnoreCase));
                for (int i = 0; i < args.Length; i++)
                {
                    if (string.Equals(args[i], "--printer", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length) { printer = args[i + 1]; i++; continue; }
                    if (string.Equals(args[i], "--template", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length) { templateId = args[i + 1]; i++; continue; }
                    if (string.Equals(args[i], "--dpi", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length && int.TryParse(args[i + 1], out var d)) { dpi = Math.Clamp(d, 150, 1200); i++; continue; }
                }
                pdfs = args.Where(a => a.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && File.Exists(a)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "解析命令行参数失败");
            }

            if (toPrint && pdfs.Length > 0)
            {
                // 静默打印模式：不创建任何窗口；先启动 Dispatcher 循环，再在后台执行打印，完成后关闭
                this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
                base.OnStartup(e); // 确保 Dispatcher 开始泵消息，便于 WinRT/Invoke 正常工作

                _ = System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        var svc = new Pdftools.Desktop.Services.PrintServiceSystem();
                        var settings = Pdftools.Core.Services.SettingsService.Load();
                        string? printerName = printer ?? settings.DefaultPrinter;
                        if (string.IsNullOrWhiteSpace(printerName))
                        {
                            var all = svc.GetPrinters();
                            if (all.Length > 0) printerName = all[0];
                        }
                        string tpl = templateId ?? settings.DefaultTemplateId ?? "default";
                        if (!string.IsNullOrWhiteSpace(printerName))
                        {
                            foreach (var f in pdfs)
                            {
                                try
                                {
                                    // 在 UI Dispatcher 上执行整段打印，确保 WinRT 渲染与 XPS Writer 都在 STA/Dispatcher 线程
                                    this.Dispatcher.Invoke(() =>
                                    {
                                        var svcLocal = new Pdftools.Desktop.Services.PrintServiceSystem();
                                        svcLocal.PrintA5ToA4Top(f, printerName!, tpl, dpi);
                                    });
                                }
                                catch (Exception ex1) { Log.Error(ex1, "Silent print failed: {File}", f); }
                            }
                        }
                        else
                        {
                            Log.Warning("静默打印：未找到可用打印机，跳过");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "静默打印失败");
                    }
                    finally
                    {
                        Log.Information("静默打印完成，退出应用");
                        try { this.Dispatcher.Invoke(() => Shutdown(0)); } catch { Shutdown(0); }
                    }
                });
                return;
            }

            // 非静默：创建窗口，并在需要时导入 PDF
            base.OnStartup(e);

            // 启动自检（仅交互模式，避免无谓开销）
            _ = SelfCheckAsync();

            var mw = new MainWindow();
            this.MainWindow = mw;
            mw.Show();
            if (pdfs.Length > 0)
            {
                try { mw.ImportPaths(pdfs); } catch (Exception ex) { Log.Error(ex, "启动导入失败"); }
            }
        }

        // 自检：验证 PDF 预览渲染（内部自捕获异常，避免打断启动）
        private async Task SelfCheckAsync()
        {
            // 1) PDF 渲染可用性：生成一页简单 PDF 到内存，然后用 Windows.Data.Pdf 渲染（全程内存，不落磁盘）
            try
            {
                using var msPdf = new MemoryStream();
                using (var doc = new PdfSharpDoc())
                {
                    var page = doc.AddPage();
                    using var gfx = XGraphics.FromPdfPage(page);
                    // 避免字体解析依赖，仅绘制图形
                    gfx.DrawRectangle(XPens.Black, new XRect(10, 10, 50, 30));
                    doc.Save(msPdf, false);
                }
                // 将托管流内容写入 WinRT 内存流（单次拷贝）
                msPdf.Position = 0;
                using var ra = new InMemoryRandomAccessStream();
                var writer = new DataWriter(ra);
                var bytesPdf = msPdf.ToArray();
                writer.WriteBytes(bytesPdf);
                await writer.StoreAsync();
                writer.DetachStream();
                writer.Dispose();
                ra.Seek(0);
                var pdf = await global::Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(ra);
                using (var page0 = pdf.GetPage(0))
                using (var outStream = new InMemoryRandomAccessStream())
                {
                    await page0.RenderToStreamAsync(outStream, new global::Windows.Data.Pdf.PdfPageRenderOptions { DestinationWidth = 100, DestinationHeight = 100 });
                    // 将 WinRT 流转换为托管内存
                    outStream.Seek(0);
                    var reader = new DataReader(outStream.GetInputStreamAt(0));
                    uint size = (uint)outStream.Size;
                    await reader.LoadAsync(size);
                    var bytes = new byte[size];
                    reader.ReadBytes(bytes);
                    reader.Dispose();
                    using var msOut = new MemoryStream(bytes);
                    var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(msOut, System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat, System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                    var frame = decoder.Frames[0];
                    if (frame.PixelWidth > 0) Log.Information("SelfCheck: PDF 渲染正常"); else throw new InvalidOperationException("Render empty");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "SelfCheck: PDF 渲染失败");
                // 自检失败不阻塞启动，仅记录日志，不弹窗打断启动流程
            }

            // 压缩功能已移除
        }


        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("应用退出");
            Log.CloseAndFlush();
            base.OnExit(e);
        }
    }
}
