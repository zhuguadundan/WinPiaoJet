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
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            LoadPrintersSafe();
            LoadTemplates();
            // 加载设置
            var settings = Pdftools.Core.Services.SettingsService.Load();
            if (!string.IsNullOrEmpty(settings.DefaultPrinter))
            {
                var printers = PrinterCombo.ItemsSource as string[];
                if (printers != null)
                {
                    var idx = Array.IndexOf(printers, settings.DefaultPrinter);
                    if (idx >= 0) PrinterCombo.SelectedIndex = idx;
                }
            }
            if (!string.IsNullOrEmpty(settings.DefaultTemplateId))
            {
                TemplateCombo.SelectedValue = settings.DefaultTemplateId;
            }
            // 默认选择边距 3mm 与安全距 5mm
            MarginCombo.SelectedIndex = 1;
            SafeGapCombo.SelectedIndex = 0;
            // 预览/打印 DPI 默认
            PreviewDpiCombo.SelectedIndex = 0; // 150
            PrintDpiCombo.SelectedIndex = 1;   // 300
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

        private void MarginCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

        private void SafeGapCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

        private void AutoShrinkCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

        private void PreviewDpiCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

        private void LoadTemplates()
        {
            var svc = new Pdftools.Core.Services.TemplateService();
            var all = svc.GetAll();
            TemplateCombo.ItemsSource = all;
            TemplateCombo.DisplayMemberPath = nameof(Pdftools.Core.Models.Template.Name);
            TemplateCombo.SelectedValuePath = nameof(Pdftools.Core.Models.Template.Id);
            // 选中 default
            var def = all.FirstOrDefault(t => t.Id == "default");
            if (def != null)
            {
                TemplateCombo.SelectedValue = def.Id;
                ApplyTemplateToUi(def);
            }
        }

        private void TemplateCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (TemplateCombo.SelectedValue is string id)
            {
                _currentTemplateId = id;
                var svc = new Pdftools.Core.Services.TemplateService();
                var tpl = svc.Get(id);
                ApplyTemplateToUi(tpl);
                // 保存默认模板到设置
                var settings = Pdftools.Core.Services.SettingsService.Load();
                settings.DefaultTemplateId = id;
                Pdftools.Core.Services.SettingsService.Save(settings);
                if (!string.IsNullOrEmpty(_currentFilePath)) _ = UpdatePlacementPreviewAsync();
            }
        }

        private void ApplyTemplateToUi(Pdftools.Core.Models.Template tpl)
        {
            // 边距
            var idx = tpl.MarginMm switch { 0 => 0, 3 => 1, 5 => 2, 8 => 3, _ => 1 };
            MarginCombo.SelectedIndex = idx;
            AutoShrinkCheck.IsChecked = tpl.AutoShrink;
            // 安全距
            var sidx = tpl.SafeGapMm switch { 5 => 0, 7 => 1, 10 => 2, _ => 0 };
            SafeGapCombo.SelectedIndex = sidx;
        }

        private void BtnSaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            var svc = new Pdftools.Core.Services.TemplateService();
            var name = $"自定义模板 {DateTime.Now:HHmmss}";
            var id = $"tpl-{DateTime.Now:yyyyMMddHHmmss}";
            var newTpl = new Pdftools.Core.Models.Template
            {
                Id = id,
                Name = name,
                MarginMm = GetSelectedMarginMm(),
                AutoShrink = AutoShrinkCheck.IsChecked == true,
                SafeGapMm = 5
            };
            svc.Save(newTpl);
            LoadTemplates();
            TemplateCombo.SelectedValue = id;
        }
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
            int printDpi = GetSelectedPrintDpi();
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
            int marginMm = GetSelectedMarginMm();
            int safeGapMm = GetSelectedSafeGapMm();
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
            int printDpi = GetSelectedPrintDpi();
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
                int printDpi = GetSelectedPrintDpi();
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


        private int GetSelectedMarginMm()
        {
            if (MarginCombo.SelectedItem is System.Windows.Controls.ComboBoxItem cbi && int.TryParse(cbi.Content?.ToString(), out var mm))
                return mm;
            return 3;
        }

        private int GetSelectedSafeGapMm()
        {
            if (int.TryParse(SafeGapCustom.Text, out var custom) && custom > 0 && custom <= 30)
                return custom;
            if (SafeGapCombo.SelectedItem is System.Windows.Controls.ComboBoxItem cbi && int.TryParse(cbi.Content?.ToString(), out var mm))
                return mm;
            return 5;
        }

        private void SafeGapCustom_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentFilePath))
                _ = UpdatePlacementPreviewAsync();
        }

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
                int marginMm = GetSelectedMarginMm();
                int safeGapMm = GetSelectedSafeGapMm();
                bool autoShrink = AutoShrinkCheck.IsChecked == true;
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
                    int dpiPreview = GetSelectedPreviewDpi();
                    var cacheKey = $"{_currentFilePath}|0|{dpiPreview}|{System.IO.File.GetLastWriteTimeUtc(_currentFilePath).Ticks}";
                    if (!_previewCache.TryGetValue(cacheKey, out var bs))
                    {
                        bs = await Pdftools.Desktop.Services.WindowsPdfRenderer.RenderPageAsync(_currentFilePath, 0, dpiPreview);
                        _previewCache[cacheKey] = bs;
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

                var tplSvc = new Pdftools.Core.Services.TemplateService();
                var tpl = tplSvc.Get(_currentTemplateId);
                tpl.MarginMm = marginMm; tpl.AutoShrink = autoShrink; tpl.SafeGapMm = safeGapMm; tplSvc.Save(tpl);

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

        private int GetSelectedPreviewDpi()
        {
            if (PreviewDpiCombo.SelectedItem is System.Windows.Controls.ComboBoxItem cbi && int.TryParse(cbi.Content?.ToString(), out var dpi)) return dpi;
            return 150;
        }

        private int GetSelectedPrintDpi()
        {
            if (PrintDpiCombo.SelectedItem is System.Windows.Controls.ComboBoxItem cbi && int.TryParse(cbi.Content?.ToString(), out var dpi)) return dpi;
            return 300;
        }
    }
}
