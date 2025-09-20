using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Threading.Tasks;
using PdfSharp.Pdf.IO;
using Pdftools.Core.Services;
using Serilog;

namespace Pdftools.Desktop.Services;

/// <summary>
/// 基于 System.Printing 的打印服务：
/// - GetPrinters：枚举打印机
/// - PrintA5ToA4Top：位图回退路径，按“上半张 1:1 + 自动微缩”绘制
///   注意：向量直打将在后续替换；当前先保证正确位置与比例。
/// </summary>
public class PrintServiceSystem : IPrintService
{
    private const double MmPerInch = 25.4;
    private const double DipPerInch = 96.0; // WPF 设备无关像素
    private const double PtPerInch = 72.0;

    public string[] GetPrinters()
    {
        var list = new List<string>();
        using var server = new LocalPrintServer();
        var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
        foreach (var q in queues)
        {
            try { list.Add(q.FullName); } catch { /* ignore */ }
        }
        return list.OrderBy(s => s).ToArray();
    }

    public void PrintA5ToA4Top(string filePath, string printerName, string templateId)
    {
        Log.Information("PrintA5ToA4Top: start file={File} printer={Printer} tpl={Tpl}", filePath, printerName, templateId);
        // 读取模板参数
        var templateSvc = new Pdftools.Core.Services.TemplateService();
        var tpl = templateSvc.Get(templateId);
        int marginMm = tpl.MarginMm;
        int safeGapMm = tpl.SafeGapMm;
        bool autoShrink = tpl.AutoShrink;

        // 读取 PDF 页面尺寸（pt）——使用 PDFsharp，仅用于尺寸，不涉及渲染
        double pageWpt, pageHpt;
        using (var src = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
        {
            var p = src.Pages[0];
            pageWpt = p.Width.Point; pageHpt = p.Height.Point;
        }

        // 获取打印队列与能力
        using var server = new LocalPrintServer();
        var queue = server.GetPrintQueue(printerName);
        var ticket = queue.DefaultPrintTicket ?? new PrintTicket();
        ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
        var caps = queue.GetPrintCapabilities(ticket);
        var area = caps.PageImageableArea; // DIP (1/96")
        if (area == null) throw new InvalidOperationException("无法读取打印机可打印区域");

        double printableWidthDip = area.ExtentWidth;
        double printableHeightDip = area.ExtentHeight;
        double originXDip = area.OriginWidth;
        double originYDip = area.OriginHeight;

        // 上半区域（DIP）与布局计算（复用 Core 算法以保持一致）
        var layout = Pdftools.Core.Services.PrintLayout.ComputeA5Top(
            pageWpt, pageHpt,
            printableWidthMm: printableWidthDip / DipPerInch * MmPerInch,
            printableHeightMm: printableHeightDip / DipPerInch * MmPerInch,
            marginMm: marginMm,
            safeGapMm: safeGapMm,
            autoShrink: autoShrink);
        Log.Information("Print layout: content={W}x{H}mm offset=({X},{Y})mm scale={Scale:0.###}",
            layout.ContentWidthMm, layout.ContentHeightMm, layout.OffsetLeftMm, layout.OffsetTopMm, layout.Scale);

        double contentWdip = layout.ContentWidthMm / MmPerInch * DipPerInch;
        double contentHdip = layout.ContentHeightMm / MmPerInch * DipPerInch;
        double offsetX = originXDip + layout.OffsetLeftMm / MmPerInch * DipPerInch;
        double offsetY = originYDip + layout.OffsetTopMm / MmPerInch * DipPerInch; // 左上锚点

        // 渲染位图（高 DPI，使用 Windows.Data.Pdf）
        int dpi = 600;
        var image = Pdftools.Desktop.Services.WindowsPdfRenderer.RenderPageAsync(filePath, 0, dpi)
            .GetAwaiter().GetResult();

        // 绘制到 Visual 并发送到打印机（直接使用 WPF BitmapSource）
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            // 将图像绘制到指定矩形
            dc.DrawImage(image, new Rect(offsetX, offsetY, contentWdip, contentHdip));
        }

        var writer = PrintQueue.CreateXpsDocumentWriter(queue);
        Log.Information("Sending to printer {Printer}...", printerName);
        writer.Write(visual, ticket);
        Log.Information("PrintA5ToA4Top: completed file={File}", filePath);
    }

    public void PrintCalibrationPage(string printerName, int marginMm, int safeGapMm)
    {
        using var server = new LocalPrintServer();
        var queue = server.GetPrintQueue(printerName);
        var ticket = queue.DefaultPrintTicket ?? new PrintTicket();
        ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
        var caps = queue.GetPrintCapabilities(ticket);
        var area = caps.PageImageableArea ?? throw new InvalidOperationException("无法读取打印机可打印区域");

        double pw = area.ExtentWidth;
        double ph = area.ExtentHeight;
        double ox = area.OriginWidth;
        double oy = area.OriginHeight;

        double mmToDip = DipPerInch / MmPerInch;
        double marginDip = marginMm * mmToDip;
        double safeDip = safeGapMm * mmToDip;

        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            dc.PushTransform(new TranslateTransform(ox, oy));
            // 打印区域背景
            dc.DrawRectangle(System.Windows.Media.Brushes.White, new System.Windows.Media.Pen(System.Windows.Media.Brushes.Gray, 1), new Rect(0, 0, pw, ph));

            // 标尺
            var tickPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Black, 1);
            var thinPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Gray, 0.8);
            var tf = new System.Windows.Media.Typeface("Segoe UI");
            for (int mm = 0; mm <= (int)(pw / mmToDip); mm += 5)
            {
                double x = mm * mmToDip;
                double len = (mm % 10 == 0) ? 8 : 4;
                dc.DrawLine(mm % 10 == 0 ? tickPen : thinPen, new System.Windows.Point(x, 0), new System.Windows.Point(x, len));
                if (mm % 10 == 0)
                {
                    var ft = new System.Windows.Media.FormattedText(mm.ToString(), System.Globalization.CultureInfo.CurrentUICulture,
                        System.Windows.FlowDirection.LeftToRight, tf, 10, System.Windows.Media.Brushes.Black, 1.0);
                    dc.DrawText(ft, new System.Windows.Point(x + 2, len + 1));
                }
            }
            for (int mm = 0; mm <= (int)(ph / mmToDip); mm += 5)
            {
                double y = mm * mmToDip;
                double len = (mm % 10 == 0) ? 8 : 4;
                dc.DrawLine(mm % 10 == 0 ? tickPen : thinPen, new System.Windows.Point(0, y), new System.Windows.Point(len, y));
                if (mm % 10 == 0)
                {
                    var ft = new System.Windows.Media.FormattedText(mm.ToString(), System.Globalization.CultureInfo.CurrentUICulture,
                        System.Windows.FlowDirection.LeftToRight, tf, 10, System.Windows.Media.Brushes.Black, 1.0);
                    // 旋转文本写在左边
                    dc.PushTransform(new RotateTransform(-90, 0, 0));
                    dc.DrawText(ft, new System.Windows.Point(-y - ft.Width - 2, len + 14));
                    dc.Pop();
                }
            }

            // 上半分割线与安全线
            var midPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.LightGray, 1);
            dc.DrawLine(midPen, new System.Windows.Point(0, ph / 2.0), new System.Windows.Point(pw, ph / 2.0));
            var dash = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Orange, 1) { DashStyle = System.Windows.Media.DashStyles.Dash };
            dc.DrawLine(dash, new System.Windows.Point(0, ph / 2.0 - safeDip), new System.Windows.Point(pw, ph / 2.0 - safeDip));

            // 边距框（上半区域内容边框）
            var contentRect = new Rect(marginDip, marginDip, pw - marginDip * 2, ph / 2.0 - safeDip - marginDip);
            dc.DrawRectangle(null, new System.Windows.Media.Pen(System.Windows.Media.Brushes.Blue, 1), contentRect);

            // 标注信息
            var note = new System.Windows.Media.FormattedText($"Margin={marginMm}mm  Safe={safeGapMm}mm", System.Globalization.CultureInfo.CurrentUICulture,
                System.Windows.FlowDirection.LeftToRight, tf, 12, System.Windows.Media.Brushes.Black, 1.0);
            dc.DrawText(note, new System.Windows.Point(10, ph - 24));

            dc.Pop();
        }

        var writer = PrintQueue.CreateXpsDocumentWriter(queue);
        writer.Write(visual, ticket);
    }
}
