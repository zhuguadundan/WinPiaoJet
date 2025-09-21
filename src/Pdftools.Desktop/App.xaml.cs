using System;
using System.IO;
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

            // 启动自检（全程内存管道，不落磁盘；不阻塞启动）
            _ = SelfCheckAsync();

            Log.Information("应用启动");
            base.OnStartup(e);
        }

        // 自检：验证 PDF 预览渲染与 qpdf 可用性（内部自捕获异常，避免打断启动）
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
                // 将托管流内容写入 WinRT 内存流，避免使用扩展方法
                msPdf.Position = 0;
                var ra = new InMemoryRandomAccessStream();
                using (var temp = new MemoryStream(msPdf.ToArray()))
                {
                    var writer = new DataWriter(ra);
                    writer.WriteBytes(temp.ToArray());
                    await writer.StoreAsync();
                    writer.DetachStream();
                    writer.Dispose();
                }
                ra.Seek(0);
                var pdf = await global::Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(ra);
                using var page0 = pdf.GetPage(0);
                using var outStream = new InMemoryRandomAccessStream();
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
            catch (Exception ex)
            {
                Log.Error(ex, "SelfCheck: PDF 渲染失败");
                // 自检失败不阻塞启动，仅记录日志，不弹窗打断启动流程
            }

            // 2) qpdf 可用性（可选）
            try
            {
                bool ok = QpdfRunner.IsAvailable();
                Log.Information("SelfCheck: qpdf 可用性 = {Available}", ok);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "SelfCheck: 检测 qpdf 失败");
            }
        }


        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("应用退出");
            Log.CloseAndFlush();
            base.OnExit(e);
        }
    }
}
