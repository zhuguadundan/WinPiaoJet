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

    public void PrintA5ToA4Top(string filePath, string printerName, string templateId, int dpi = 300)
    {
        Log.Information("PrintA5ToA4Top: start file={File} printer={Printer} tpl={Tpl}", filePath, printerName, templateId);
        // 读取模板参数
        var templateSvc = new Pdftools.Core.Services.TemplateService();
        var tpl = templateSvc.Get(templateId);
        int marginMm = tpl.MarginMm;
        int safeGapMm = tpl.SafeGapMm;
        bool autoShrink = tpl.AutoShrink; // 取消“强制不缩放”，仅由模板的 AutoShrink 决定

        // 读取 PDF 页面尺寸（pt）——使用 PDFsharp，仅用于尺寸，不涉及渲染
        double pageWpt, pageHpt;
        using (var src = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
        {
            var p = src.Pages[0];
            pageWpt = p.Width.Point; pageHpt = p.Height.Point;
        }

        // 基于页面尺寸动态限制渲染像素，避免过大位图导致驱动失败/内存峰值过高
        double pageWdip = pageWpt / PtPerInch * DipPerInch;
        double pageHdip = pageHpt / PtPerInch * DipPerInch;
        double pxW = pageWdip * dpi / DipPerInch;
        double pxH = pageHdip * dpi / DipPerInch;
        const double MaxPixels = 12_000_000; // 约 12MP，更保守以兼容部分驱动
        double pixels = pxW * pxH;
        if (pixels > MaxPixels)
        {
            double s = Math.Sqrt(MaxPixels / pixels);
            int newDpi = (int)Math.Max(150, Math.Floor(dpi * s));
            Log.Warning("Print render dpi adjusted from {OldDpi} to {NewDpi} due to pixel cap (page {W}x{H}dip)", dpi, newDpi, pageWdip, pageHdip);
            dpi = newDpi;
        }

        // 预先在后台线程获取打印能力并计算布局与目标像素尺寸，然后离屏渲染为 PNG 字节
        using var serverPre = new LocalPrintServer();
        using var queuePre = serverPre.GetPrintQueue(printerName);
        var ticketPre = queuePre.DefaultPrintTicket ?? new PrintTicket();
        ticketPre.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);
        var validatedPre = queuePre.MergeAndValidatePrintTicket(queuePre.DefaultPrintTicket, ticketPre);
        var vtPre = validatedPre.ValidatedPrintTicket ?? ticketPre;
        var capsPre = queuePre.GetPrintCapabilities(vtPre);
        var areaPre = capsPre.PageImageableArea ?? throw new InvalidOperationException("无法读取打印机可打印区域");
        double printableWidthDipPre = areaPre.ExtentWidth;
        double printableHeightDipPre = areaPre.ExtentHeight;

        var layoutPre = Pdftools.Core.Services.PrintLayout.ComputeA5Top(
            pageWpt, pageHpt,
            printableWidthMm: printableWidthDipPre / DipPerInch * MmPerInch,
            printableHeightMm: printableHeightDipPre / DipPerInch * MmPerInch,
            marginMm: marginMm,
            safeGapMm: safeGapMm,
            autoShrink: autoShrink);
        double contentWdipPre = layoutPre.ContentWidthMm / MmPerInch * DipPerInch;
        double contentHdipPre = layoutPre.ContentHeightMm / MmPerInch * DipPerInch;
        int dstWpxPre = Math.Max(1, (int)Math.Round(contentWdipPre * dpi / DipPerInch));
        int dstHpxPre = Math.Max(1, (int)Math.Round(contentHdipPre * dpi / DipPerInch));
        double pixelsTargetPre = (double)dstWpxPre * dstHpxPre;
        const double MaxPixelsLocal = 12_000_000;
        if (pixelsTargetPre > MaxPixelsLocal)
        {
            double s = Math.Sqrt(MaxPixelsLocal / pixelsTargetPre);
            int dpiAdj = (int)Math.Max(150, Math.Floor(dpi * s));
            Log.Warning("Adjust print dpi from {Old} to {New} due to target pixel cap", dpi, dpiAdj);
            dpi = dpiAdj;
            dstWpxPre = Math.Max(1, (int)Math.Round(contentWdipPre * dpi / DipPerInch));
            dstHpxPre = Math.Max(1, (int)Math.Round(contentHdipPre * dpi / DipPerInch));
        }

        Log.Debug("PreRender: render PDF to {W}x{H} pixels off UI thread...", dstWpxPre, dstHpxPre);
        var (pngBytes, pxWOut, pxHOut) = Pdftools.Desktop.Services.WindowsPdfRenderer
            .RenderPagePngAsync(filePath, 0, dstWpxPre, dstHpxPre)
            .GetAwaiter().GetResult();

        void DoPrint()
        {
            try
            {
                Log.Debug("DoPrint: acquiring queue and ticket...");
                using var server = new LocalPrintServer();
                using var queue = server.GetPrintQueue(printerName);
                var ticket = queue.DefaultPrintTicket ?? new PrintTicket();
                ticket.PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4);

                Log.Debug("DoPrint: merging/validating ticket...");
                var validated = queue.MergeAndValidatePrintTicket(queue.DefaultPrintTicket, ticket);
                var vt = validated.ValidatedPrintTicket ?? ticket;

                Log.Debug("DoPrint: capabilities & area...");
                var caps = queue.GetPrintCapabilities(vt);
                var area = caps.PageImageableArea; // DIP (1/96")
                if (area == null) throw new InvalidOperationException("无法读取打印机可打印区域");

                // 基于预计算的布局，使用 PNG 字节在 UI 线程解码
                var printableWidthDip = area.ExtentWidth;
                var printableHeightDip = area.ExtentHeight;
                var originXDip = area.OriginWidth;
                var originYDip = area.OriginHeight;

                var layout = Pdftools.Core.Services.PrintLayout.ComputeA5Top(
                    pageWpt, pageHpt,
                    printableWidthMm: printableWidthDip / DipPerInch * MmPerInch,
                    printableHeightMm: printableHeightDip / DipPerInch * MmPerInch,
                    marginMm: marginMm,
                    safeGapMm: safeGapMm,
                    autoShrink: autoShrink);

                double contentWdip = layout.ContentWidthMm / MmPerInch * DipPerInch;
                double contentHdip = layout.ContentHeightMm / MmPerInch * DipPerInch;
                double offsetX = originXDip + layout.OffsetLeftMm / MmPerInch * DipPerInch;
                double offsetY = originYDip + layout.OffsetTopMm / MmPerInch * DipPerInch; // 左上锚点

                Log.Debug("DoPrint: building RTB from PNG bytes (use actual decoded size, 1:1 mapping)...");
                using var ms = new MemoryStream(pngBytes);
                var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                var frame = decoder.Frames[0];
                int fw = frame.PixelWidth;
                int fh = frame.PixelHeight;
                double fwDip = fw * 96.0 / dpi;
                double fhDip = fh * 96.0 / dpi;
                var dvCopy = new DrawingVisual();
                using (var dcc = dvCopy.RenderOpen())
                {
                    dcc.DrawRectangle(System.Windows.Media.Brushes.White, null, new Rect(0, 0, fwDip, fhDip));
                    dcc.DrawImage(frame, new Rect(0, 0, fwDip, fhDip));
                }
                var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(
                    fw, fh, dpi, dpi, System.Windows.Media.PixelFormats.Pbgra32);
                System.Windows.Media.RenderOptions.SetBitmapScalingMode(dvCopy, System.Windows.Media.BitmapScalingMode.NearestNeighbor);
                rtb.Render(dvCopy);
                if (rtb.CanFreeze) rtb.Freeze();

                Log.Debug("DoPrint: prepare visual...");
                var visual = new DrawingVisual();
                System.Windows.Media.RenderOptions.SetBitmapScalingMode(visual, System.Windows.Media.BitmapScalingMode.NearestNeighbor);
                using (var dc = visual.RenderOpen())
                {
                    dc.DrawImage(rtb, new Rect(offsetX, offsetY, contentWdip, contentHdip));
                }

                Log.Debug("DoPrint: create writer & send...");
                var writer = PrintQueue.CreateXpsDocumentWriter(queue);
                Log.Information("Sending to printer {Printer}... (ticket validated)", printerName);
                try
                {
                    Log.Debug("before-writer.Write: rtbFrozen={Frozen}", rtb.IsFrozen);
                    writer.Write(visual, vt);
                    Log.Information("PrintA5ToA4Top: completed file={File}", filePath);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "PrintA5ToA4Top: writer.Write 失败 (first try)");
                    try
                    {
                        var converted24 = new System.Windows.Media.Imaging.FormatConvertedBitmap(
                            rtb, System.Windows.Media.PixelFormats.Bgr24, null, 0);
                        if (converted24.CanFreeze) converted24.Freeze();
                        var visual2 = new DrawingVisual();
                        System.Windows.Media.RenderOptions.SetBitmapScalingMode(visual2, System.Windows.Media.BitmapScalingMode.NearestNeighbor);
                        using (var dc2 = visual2.RenderOpen())
                        {
                            dc2.DrawImage(converted24, new Rect(offsetX, offsetY, contentWdip, contentHdip));
                        }
                        Log.Warning("Retry writer.Write with 24bpp Bgr24");
                        writer.Write(visual2, vt);
                        Log.Information("PrintA5ToA4Top: completed after retry file={File}", filePath);
                    }
                    catch (Exception ex2)
                    {
                        Log.Error(ex2, "PrintA5ToA4Top: writer.Write 失败 (retry)");
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "DoPrint failed before writer.Write");
                throw;
            }
        }

        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(DoPrint);
        }
        else
        {
            DoPrint();
        }
    }

    public void PrintCalibrationPage(string printerName, int marginMm, int safeGapMm)
    {
        using var server = new LocalPrintServer();
        using var queue = server.GetPrintQueue(printerName);
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
