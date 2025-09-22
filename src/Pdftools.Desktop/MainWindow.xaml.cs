using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Pdftools.Desktop.Models;
using System.Linq;

namespace Pdftools.Desktop
{
    /// <summary>
    /// 主窗口：加入简单的 PDF 打开与第一页预览
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? _currentFilePath;
        public ObservableCollection<FileItem> Files { get; } = new();
        private string _currentTemplateId = "default";
        // 预览位图缓存：(file|page|dpi|mtimeTicks) -> BitmapSource
        private readonly Dictionary<string, BitmapSource> _previewCache = new();
        // 简易 LRU 队列，限制缓存容量，防止内存涨高
        private readonly System.Collections.Generic.LinkedList<string> _previewCacheOrder = new();
        private const int PreviewCacheCapacity = 20;
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            LoadPrintersSafe();
            // 记忆默认打印机
            var settings = Pdftools.Core.Services.SettingsService.Load();
            if (!string.IsNullOrEmpty(settings.DefaultPrinter))
            {
                if (PrinterCombo.ItemsSource is string[] printers)
                {
                    var idx = Array.IndexOf(printers, settings.DefaultPrinter);
                    if (idx >= 0) PrinterCombo.SelectedIndex = idx;
                }
            }
        }

        private void LoadPrintersSafe()
        {
            try
            {
                var svc = new Services.PrintServiceSystem();
                var printers = svc.GetPrinters();
                PrinterCombo.ItemsSource = printers;
                if (printers.Length > 0) PrinterCombo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                StatusText.Content = $"读取打印机失败：{ex.Message}";
            }
        }

        private void PrinterCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

        // 模板相关操作移至“设置”菜单
        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PDF 文件 (*.pdf)|*.pdf",
                Multiselect = true
            };
            if (ofd.ShowDialog() == true)
            {
                AddFiles(ofd.FileNames);
            }
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            using var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var files = Directory.EnumerateFiles(dlg.SelectedPath, "*.pdf", SearchOption.TopDirectoryOnly).ToArray();
                AddFiles(files);
            }
        }

        private void AddFiles(string[] paths)
        {
            foreach (var p in paths)
            {
                try
                {
                    int pages = 0; double w=0, h=0;
                    try
                    {
                        using var src = PdfSharp.Pdf.IO.PdfReader.Open(p, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                        pages = src.PageCount;
                        var pg = src.Pages[0];
                        w = pg.Width.Point; h = pg.Height.Point;
                    }
                    catch { }
                    string sizeTxt = pages > 0 ? $"{w:0}x{h:0} pt" : "-";
                    Files.Add(new FileItem
                    {
                        Name = System.IO.Path.GetFileName(p),
                        Path = p,
                        PageCount = pages,
                        Size = sizeTxt,
                        Status = "待打印",
                        Message = string.Empty
                    });
                }
                catch (Exception ex)
                {
                    Files.Add(new FileItem
                    {
                        Name = System.IO.Path.GetFileName(p),
                        Path = p,
                        PageCount = 0,
                        Size = "-",
                        Status = "失败",
                        Message = ex.Message
                    });
                }
            }
            if (paths.Length > 0)
            {
                _currentFilePath = paths[0];
                LoadAndPreview(_currentFilePath);
            }
        }

        private void FilesGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is FileItem fi)
            {
                _currentFilePath = fi.Path;
                LoadAndPreview(_currentFilePath);
            }
        }

        /// <summary>
        /// 启动参数导入入口：供 App 在 OnStartup 后调用
        /// </summary>
        /// <param name="paths">PDF 路径（绝对路径）</param>
        public void ImportPaths(string[] paths)
        {
            if (paths == null || paths.Length == 0) return;
            var pdfs = paths.Where(p => !string.IsNullOrWhiteSpace(p) && System.IO.Path.GetExtension(p).Equals(".pdf", StringComparison.OrdinalIgnoreCase) && File.Exists(p)).ToArray();
            if (pdfs.Length == 0) return;
            AddFiles(pdfs);
        }

        private void FilesGrid_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            {
                var dropped = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
                var pdfs = dropped.Where(f => string.Equals(System.IO.Path.GetExtension(f), ".pdf", StringComparison.OrdinalIgnoreCase)).ToArray();
                if (pdfs.Length > 0) AddFiles(pdfs);
            }
        }

        private void BtnPrintAll_Click(object sender, RoutedEventArgs e)
        {
            if (_batchTask != null && !_batchTask.IsCompleted)
            {
                StatusText.Content = "批量任务已在运行";
                return;
            }
            if (PrinterCombo.SelectedItem is not string printer)
            {
                StatusText.Content = "请先选择打印机";
                return;
            }
            if (Files.Count == 0)
            {
                StatusText.Content = "列表为空";
                return;
            }
            _cts = new System.Threading.CancellationTokenSource();
            var token = _cts.Token;
            _resumeEvent.Set();
            StatusText.Content = "开始批量打印...";
            var svc = new Services.PrintServiceSystem();
            int printDpi = 300; // 统一默认打印 DPI
            _batchTask = System.Threading.Tasks.Task.Run(() =>
            {
                int ok = 0, fail = 0, total = Files.Count;
                int processed = 0;
                foreach (var fi in Files)
                {
                    if (token.IsCancellationRequested) break;
                    _resumeEvent.Wait(token);
                    try
                    {
                        if (!File.Exists(fi.Path)) throw new FileNotFoundException("文件不存在", fi.Path);
                        System.Windows.Application.Current.Dispatcher.Invoke(() => { fi.Status = "打印中"; fi.Message = string.Empty; });
                        svc.PrintA5ToA4Top(fi.Path, printer, _currentTemplateId, printDpi);
                        System.Windows.Application.Current.Dispatcher.Invoke(() => { fi.Status = "成功"; fi.Message = string.Empty; });
                        ok++;
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Application.Current.Dispatcher.Invoke(() => { fi.Status = "失败"; fi.Message = ex.Message; });
                        fail++;
                    }
                    processed++;
                    int percent = (int)(processed * 100.0 / total);
                    System.Windows.Application.Current.Dispatcher.Invoke(() => { BatchProgress.Value = percent; StatusCount.Content = $"{processed}/{total}"; });
                }
                System.Windows.Application.Current.Dispatcher.Invoke(() => { StatusText.Content = $"打印完成：成功 {ok}，失败 {fail}"; });
            }, token);
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "CSV 文件 (*.csv)|*.csv", FileName = $"导出_{DateTime.Now:yyyyMMdd_HHmmss}.csv" };
            if (sfd.ShowDialog() == true)
            {
                using var sw = new StreamWriter(sfd.FileName);
                sw.WriteLine("名称,页数,尺寸,状态,结果,路径");
                foreach (var fi in Files)
                {
                    string line = string.Join(',', new[] {
                        EscapeCsv(fi.Name), fi.PageCount.ToString(), EscapeCsv(fi.Size), EscapeCsv(fi.Status), EscapeCsv(fi.Message), EscapeCsv(fi.Path)
                    });
                    sw.WriteLine(line);
                }
                StatusText.Content = "CSV 已导出";
            }
        }

        private void BtnPrintTest_Click(object sender, RoutedEventArgs e)
        {
            if (PrinterCombo.SelectedItem is not string printer)
            {
                StatusText.Content = "请先选择打印机";
                return;
            }
            var tplSvcCal = new Pdftools.Core.Services.TemplateService();
            var tplCal = tplSvcCal.Get(_currentTemplateId);
            int marginMm = tplCal.MarginMm;
            int safeGapMm = tplCal.SafeGapMm;
            try
            {
                var svc = new Services.PrintServiceSystem();
                svc.PrintCalibrationPage(printer, marginMm, safeGapMm);
                StatusText.Content = "测试页已发送到打印机";
            }
            catch (Exception ex)
            {
                StatusText.Content = $"测试页打印失败：{ex.Message}";
            }
        }

        private System.Threading.CancellationTokenSource? _cts;
        private System.Threading.Tasks.Task? _batchTask;
        private readonly System.Threading.ManualResetEventSlim _resumeEvent = new(true);

        private void BtnPause_Click(object sender, RoutedEventArgs e)
        {
            _resumeEvent.Reset();
            StatusText.Content = "已暂停";
        }

        private void BtnResume_Click(object sender, RoutedEventArgs e)
        {
            _resumeEvent.Set();
            StatusText.Content = "已继续";
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            _resumeEvent.Set();
            StatusText.Content = "已取消";
        }

        private void BtnRetryFailed_Click(object sender, RoutedEventArgs e)
        {
            var failed = Files.Where(f => f.Status == "失败").Select(f => f.Path).ToArray();
            if (failed.Length == 0) { StatusText.Content = "无失败项"; return; }
            if (PrinterCombo.SelectedItem is not string printer) { StatusText.Content = "请先选择打印机"; return; }
            var svc = new Services.PrintServiceSystem();
            int printDpi = 300; // 默认打印 DPI
            StatusText.Content = "开始后台重试失败项...";
            _ = System.Threading.Tasks.Task.Run(() =>
            {
                int ok = 0, fail = 0; int processed = 0; int total = failed.Length;
                foreach (var p in failed)
                {
                    try { svc.PrintA5ToA4Top(p, printer, _currentTemplateId, printDpi); ok++; }
                    catch { fail++; }
                    processed++;
                    int percent = (int)(processed * 100.0 / Math.Max(1, total));
                    System.Windows.Application.Current.Dispatcher.Invoke(() => { BatchProgress.Value = percent; StatusCount.Content = $"{processed}/{total}"; });
                }
                System.Windows.Application.Current.Dispatcher.Invoke(() => { StatusText.Content = $"重试完成：成功 {ok}，失败 {fail}"; });
            });
        }

        private static string EscapeCsv(string? s)
        {
            s ??= string.Empty;
            if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
            {
                s = s.Replace("\"", "\"\"");
                return $"\"{s}\"";
            }
            return s;
        }

        private void BtnOpenTools_Click(object sender, RoutedEventArgs e)
        {
            var win = new Windows.ToolsWindow();
            win.Owner = this;
            win.Show();
        }

        private void BtnTemplates_Click(object sender, RoutedEventArgs e)
        {
            var win = new Windows.Templates.TemplatesWindow();
            win.Owner = this;
            win.ShowDialog();
        }

        private async void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                StatusText.Content = "请先导入 PDF 文件";
                return;
            }
            if (PrinterCombo.SelectedItem is not string printer)
            {
                StatusText.Content = "请先选择打印机";
                return;
            }

            StatusText.Content = "正在渲染并发送到打印机...";
            var svc = new Services.PrintServiceSystem();
            try
            {
                int printDpi = 300; // 默认打印 DPI
                await System.Threading.Tasks.Task.Run(() => svc.PrintA5ToA4Top(_currentFilePath!, printer, templateId: _currentTemplateId, dpi: printDpi));
                StatusText.Content = "已发送到打印机";
            }
            catch (Exception ex)
            {
                StatusText.Content = $"打印失败：{ex.Message}";
            }
        }

        private async void LoadAndPreview(string filePath)
        {
                try
                {
                    // 使用 PDFsharp 统计页数
                    int pageCount = 0;
                    try
                    {
                        using var src = PdfSharp.Pdf.IO.PdfReader.Open(filePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                        pageCount = src.PageCount;
                    }
                    catch { }
                StatusText.Content = $"已加载: {System.IO.Path.GetFileName(filePath)}  第1/{Math.Max(1,pageCount)}页";
            }
            catch (Exception ex)
            {
                StatusText.Content = $"加载失败：{ex.Message}";
                return;
            }
            await UpdatePlacementPreviewAsync();
        }


        // 边距/安全距改为读取模板参数；自定义输入移至“设置”

        private async Task UpdatePlacementPreviewAsync()
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;
            if (PrinterCombo.SelectedItem is not string printer) return;

            try
            {
                // 打印机可打印区域
                using var server = new System.Printing.LocalPrintServer();
                var queue = server.GetPrintQueue(printer);
                var ticket = queue.DefaultPrintTicket ?? new System.Printing.PrintTicket();
                ticket.PageMediaSize = new System.Printing.PageMediaSize(System.Printing.PageMediaSizeName.ISOA4);
                var caps = queue.GetPrintCapabilities(ticket);
                var area = caps.PageImageableArea;
                if (area == null) { StatusText.Content = "无法获取打印机可打印区域"; return; }

                double pw = area.ExtentWidth;   // DIP
                double ph = area.ExtentHeight;  // DIP

                // 使用 PDFsharp 读取页面尺寸（pt）
                double pageWpt, pageHpt;
                using (var src = PdfSharp.Pdf.IO.PdfReader.Open(_currentFilePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                {
                    var p = src.Pages[0];
                    pageWpt = p.Width.Point; pageHpt = p.Height.Point;
                }

                // 计算布局
                double dipToMm = 25.4 / 96.0;
                var tplSvcPrev = new Pdftools.Core.Services.TemplateService();
                var tplPrev = tplSvcPrev.Get(_currentTemplateId);
                int marginMm = tplPrev.MarginMm;
                int safeGapMm = tplPrev.SafeGapMm;
                bool autoShrink = tplPrev.AutoShrink;
                var layout = Pdftools.Core.Services.PrintLayout.ComputeA5Top(
                    pageWpt, pageHpt,
                    printableWidthMm: pw * dipToMm,
                    printableHeightMm: ph * dipToMm,
                    marginMm: marginMm,
                    safeGapMm: safeGapMm,
                    autoShrink: autoShrink);

                // 视图缩放
                double targetViewWidthPx = 600;
                double viewScale = targetViewWidthPx / pw;
                int bmpW = Math.Max(1, (int)Math.Round(pw * viewScale));
                int bmpH = Math.Max(1, (int)Math.Round(ph * viewScale));

                var dv = new System.Windows.Media.DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    dc.PushTransform(new System.Windows.Media.ScaleTransform(viewScale, viewScale));

                    // 背景：可打印区域
                    var printRect = new System.Windows.Rect(0, 0, pw, ph);
                    dc.DrawRectangle(System.Windows.Media.Brushes.White, new System.Windows.Media.Pen(System.Windows.Media.Brushes.Gray, 1 / viewScale), printRect);

                    // 标尺与分割线
                    double mmToDip = 96.0 / 25.4;
                    var tickPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Black, 1 / viewScale);
                    var thinPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Gray, 0.5 / viewScale);
                    var typeface = new System.Windows.Media.Typeface("Segoe UI");
                    for (int mm = 0; mm <= (int)(pw / mmToDip); mm += 5)
                    {
                        double x = mm * mmToDip;
                        double len = (mm % 10 == 0) ? 6 : 3;
                        dc.DrawLine(mm % 10 == 0 ? tickPen : thinPen, new System.Windows.Point(x, 0), new System.Windows.Point(x, len));
                    }
                    for (int mm = 0; mm <= (int)(ph / mmToDip); mm += 5)
                    {
                        double y = mm * mmToDip;
                        double len = (mm % 10 == 0) ? 6 : 3;
                        dc.DrawLine(mm % 10 == 0 ? tickPen : thinPen, new System.Windows.Point(0, y), new System.Windows.Point(len, y));
                    }
                    var midPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.LightGray, 1 / viewScale);
                    dc.DrawLine(midPen, new System.Windows.Point(0, ph / 2.0), new System.Windows.Point(pw, ph / 2.0));
                    double safeLineY = ph / 2.0 - (safeGapMm / 25.4 * 96.0);
                    var dashPen = new System.Windows.Media.Pen(System.Windows.Media.Brushes.Orange, 1 / viewScale) { DashStyle = System.Windows.Media.DashStyles.Dash };
                    dc.DrawLine(dashPen, new System.Windows.Point(0, safeLineY), new System.Windows.Point(pw, safeLineY));

                    // 内容矩形（DIP）
                    double contentWdip = layout.ContentWidthMm / 25.4 * 96.0;
                    double contentHdip = layout.ContentHeightMm / 25.4 * 96.0;
                    double offsetX = layout.OffsetLeftMm / 25.4 * 96.0;
                    double offsetY = layout.OffsetTopMm / 25.4 * 96.0;

                    // 渲染第一页（Windows.Data.Pdf），加入简单缓存以减少重复渲染
                    int dpiPreview = 150; // 默认预览 DPI
                    var cacheKey = $"{_currentFilePath}|0|{dpiPreview}|{System.IO.File.GetLastWriteTimeUtc(_currentFilePath).Ticks}";
                    if (!_previewCache.TryGetValue(cacheKey, out var bs))
                    {
                        bs = await Pdftools.Desktop.Services.WindowsPdfRenderer.RenderPageAsync(_currentFilePath, 0, dpiPreview);
                        _previewCache[cacheKey] = bs;
                        // LRU 入队
                        _previewCacheOrder.Remove(cacheKey);
                        _previewCacheOrder.AddLast(cacheKey);
                        // 超出容量则淘汰最久未用项
                        if (_previewCacheOrder.Count > PreviewCacheCapacity)
                        {
                            var oldKey = _previewCacheOrder.First!.Value;
                            _previewCacheOrder.RemoveFirst();
                            _previewCache.Remove(oldKey);
                        }
                    }
                    else
                    {
                        // 命中则刷新 LRU 顺序
                        _previewCacheOrder.Remove(cacheKey);
                        _previewCacheOrder.AddLast(cacheKey);
                    }
                    var targetRect = new System.Windows.Rect(offsetX, offsetY, contentWdip, contentHdip);
                    dc.DrawImage(bs, targetRect);

                    bool overflow = (offsetY + contentHdip) > (ph / 2.0 - (safeGapMm / 25.4 * 96.0));
                    var borderBrush = overflow && !autoShrink ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.DodgerBlue;
                    dc.DrawRectangle(null, new System.Windows.Media.Pen(borderBrush, 1 / viewScale), targetRect);

                    dc.Pop();
                }

                var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(bmpW, bmpH, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                rtb.Render(dv);
                PreviewImage.Source = rtb;

                StatusInfo.Content = $"缩放: {layout.Scale * 100:0}%  尺寸: {layout.ContentWidthMm:0.0}×{layout.ContentHeightMm:0.0} mm";

                // 预览不再即时覆盖模板参数，模板管理移至“设置”
                var settings = Pdftools.Core.Services.SettingsService.Load();
                settings.DefaultTemplateId = _currentTemplateId;
                if (PrinterCombo.SelectedItem is string pSel) settings.DefaultPrinter = pSel;
                Pdftools.Core.Services.SettingsService.Save(settings);
            }
            catch (Exception ex)
            {
                StatusText.Content = $"预览失败：{ex.Message}";
            }
        }

        // 预览/打印 DPI 改为默认值，详细参数放到“设置”

        private async void BtnImagesToPdf_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "图片|*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff", Multiselect = true };
            if (ofd.ShowDialog() == true && ofd.FileNames?.Length > 0)
            {
                var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"图片合并_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
                if (sfd.ShowDialog() == true)
                {
                    var images = ofd.FileNames;
                    var output = sfd.FileName;
                    try
                    {
                        // 在后台线程执行转换，避免阻塞 UI 线程
                        await System.Threading.Tasks.Task.Run(() =>
                            Pdftools.Core.PdfOps.ImageToPdfService.Convert(images, output, marginMm: 5, pageSize: PdfSharp.PageSize.A4)
                        );
                        StatusText.Content = "图片已合并为 PDF";
                    }
                    catch (Exception ex) { StatusText.Content = "图片转PDF失败：" + ex.Message; }
                }
            }
        }

        private async void BtnPdfToImages_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
            if (ofd.ShowDialog() == true)
            {
                var fbd = new System.Windows.Forms.FolderBrowserDialog();
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    int dpi = 300; // 默认 300 DPI
                    string fmt = "png"; // 默认 PNG
                    try
                    {
                        // 内存方式加载并逐页渲染
                        using var fsIn = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                        using var ra = new global::Windows.Storage.Streams.InMemoryRandomAccessStream();
                        var writer = new global::Windows.Storage.Streams.DataWriter(ra);
                        using (var temp = new MemoryStream())
                        {
                            await fsIn.CopyToAsync(temp);
                            var bytes = temp.ToArray();
                            writer.WriteBytes(bytes);
                            await writer.StoreAsync();
                            writer.DetachStream();
                            writer.Dispose();
                        }
                        ra.Seek(0);
                        var doc = await global::Windows.Data.Pdf.PdfDocument.LoadFromStreamAsync(ra);
                        if (doc == null || doc.PageCount == 0) { StatusText.Content = "PDF 无效"; return; }
                        for (int i = 0; i < doc.PageCount; i++)
                        {
                            using var page = doc.GetPage((uint)i);
                            var dims = page.Dimensions; // DIP
                            uint w = (uint)Math.Max(1, Math.Round(dims.MediaBox.Width * dpi / 96.0));
                            uint h = (uint)Math.Max(1, Math.Round(dims.MediaBox.Height * dpi / 96.0));
                            using var stream = new global::Windows.Storage.Streams.InMemoryRandomAccessStream();
                            var options = new global::Windows.Data.Pdf.PdfPageRenderOptions { DestinationWidth = w, DestinationHeight = h };
                            await page.RenderToStreamAsync(stream, options);

                            // 读出字节，保存为 PNG/JPG/TIFF（默认 PNG）
                            stream.Seek(0);
                            var reader = new global::Windows.Storage.Streams.DataReader(stream.GetInputStreamAt(0));
                            uint size = (uint)stream.Size;
                            await reader.LoadAsync(size);
                            var data = new byte[size];
                            reader.ReadBytes(data);
                            reader.Dispose();
                            using var ms = new MemoryStream(data);
                            var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                            var frame = decoder.Frames[0];
                            string path = System.IO.Path.Combine(fbd.SelectedPath, $"{System.IO.Path.GetFileNameWithoutExtension(ofd.FileName)}_{i + 1:D3}.{fmt}".ToLower());
                            using var fs = File.Open(path, FileMode.Create, FileAccess.Write, FileShare.None);
                            var enc = new PngBitmapEncoder();
                            enc.Frames.Add(frame);
                            enc.Save(fs);
                        }
                        StatusText.Content = "PDF 已导出为图片";
                    }
                    catch (Exception ex) { StatusText.Content = "PDF转图片失败：" + ex.Message; }
                }
            }
        }
    }
}
