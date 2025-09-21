using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using Pdftools.Core.PdfOps;
using Pdftools.Desktop.Services;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace Pdftools.Desktop.Windows;

public partial class ToolsWindow : Window
{
    private readonly List<string> _images = new();
    private string? _splitSrc;
    private string? _rotateSrc;
    private string? _watermarkSrc;
    private string? _pagenumSrc;
    private readonly System.Collections.Generic.List<string> _merges = new();
    // 压缩功能已移除

    public ToolsWindow()
    {
        InitializeComponent();
        ImgMarginCombo.SelectedIndex = 1; // 5mm
        DpiCombo.SelectedIndex = 1; // 300
        ImgFormatCombo.SelectedIndex = 0; // PNG
        ImgPageSize.SelectedIndex = 0; // A4
        PageNumberFormat.Text = "{page}/{total}";
        PageNumberPos.SelectedIndex = 4; // BottomCenter
    }

    private void BtnPickImages_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "图片|*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff", Multiselect = true };
        if (ofd.ShowDialog() == true)
        {
            _images.Clear();
            _images.AddRange(ofd.FileNames);
            ImagesList.ItemsSource = null;
            ImagesList.ItemsSource = _images;
        }
    }

    private void BtnImagesToPdf_Click(object sender, RoutedEventArgs e)
    {
        if (_images.Count == 0) { System.Windows.MessageBox.Show("请先选择图片"); return; }
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"图片合并_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            var cbi = (System.Windows.Controls.ComboBoxItem)ImgMarginCombo.SelectedItem;
            double margin = double.TryParse(cbi.Content?.ToString(), out var mm) ? mm : 5;
            try
            {
                var pageSize = ((System.Windows.Controls.ComboBoxItem)ImgPageSize.SelectedItem).Content!.ToString()! == "A5"
                    ? PdfSharp.PageSize.A5 : PdfSharp.PageSize.A4;
                ImageToPdfService.Convert(_images.ToArray(), sfd.FileName, margin, pageSize);
                System.Windows.MessageBox.Show("已导出 PDF");
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("导出失败：" + ex.Message); }
        }
    }

    private void BtnPickPdfForImages_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
        if (ofd.ShowDialog() == true)
        {
            PdfToImgTip.Text = ofd.FileName;
        }
    }

    private async void BtnPdfToImages_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(PdfToImgTip.Text) || !File.Exists(PdfToImgTip.Text)) { System.Windows.MessageBox.Show("请先选择 PDF"); return; }
        var fbd = new System.Windows.Forms.FolderBrowserDialog();
        if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            int dpi = int.Parse(((System.Windows.Controls.ComboBoxItem)DpiCombo.SelectedItem).Content!.ToString()!);
            string fmt = ((System.Windows.Controls.ComboBoxItem)ImgFormatCombo.SelectedItem).Content!.ToString()!;
            int jpegQ = 90; int.TryParse(JpegQuality.Text, out jpegQ); jpegQ = Math.Clamp(jpegQ, 50, 100);
            try
            {
                // 使用内存管道加载 PDF，避免直接占用源文件句柄
                using var fsIn = File.Open(PdfToImgTip.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                using var ra = new InMemoryRandomAccessStream();
                var writer = new DataWriter(ra);
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
                var doc = await PdfDocument.LoadFromStreamAsync(ra);
                if (doc == null || doc.PageCount == 0) { System.Windows.MessageBox.Show("PDF 无效"); return; }
                for (int i = 0; i < doc.PageCount; i++)
                {
                    using var page = doc.GetPage((uint)i);
                    var dims = page.Dimensions; // DIP (1/96")
                    double widthDip = dims.MediaBox.Width;
                    double heightDip = dims.MediaBox.Height;
                    uint w = (uint)Math.Max(1, Math.Round(widthDip * dpi / 96.0));
                    uint h = (uint)Math.Max(1, Math.Round(heightDip * dpi / 96.0));

                    using var stream = new InMemoryRandomAccessStream();
                    var options = new PdfPageRenderOptions { DestinationWidth = w, DestinationHeight = h };
                    await page.RenderToStreamAsync(stream, options);

                    // 将 WinRT 流转为 WPF BitmapFrame（不依赖扩展方法）
                    stream.Seek(0);
                    var reader = new DataReader(stream.GetInputStreamAt(0));
                    uint size = (uint)stream.Size;
                    await reader.LoadAsync(size);
                    var data = new byte[size];
                    reader.ReadBytes(data);
                    reader.Dispose();
                    using var ms = new MemoryStream(data);
                    var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    var frame = decoder.Frames[0];

                    string path = System.IO.Path.Combine(fbd.SelectedPath, $"{System.IO.Path.GetFileNameWithoutExtension(PdfToImgTip.Text)}_{i + 1:D3}.{fmt.ToLower()}");
                    using var fs = File.Open(path, FileMode.Create, FileAccess.Write, FileShare.None);
                    if (fmt.Equals("jpg", StringComparison.OrdinalIgnoreCase) || fmt.Equals("jpeg", StringComparison.OrdinalIgnoreCase))
                    {
                        var enc = new JpegBitmapEncoder { QualityLevel = jpegQ };
                        enc.Frames.Add(frame);
                        enc.Save(fs);
                    }
                    else if (fmt.Equals("tiff", StringComparison.OrdinalIgnoreCase))
                    {
                        var enc = new TiffBitmapEncoder();
                        enc.Frames.Add(frame);
                        enc.Save(fs);
                    }
                    else // png
                    {
                        var enc = new PngBitmapEncoder();
                        enc.Frames.Add(frame);
                        enc.Save(fs);
                    }
                }
                // 通过流加载已避免文件句柄占用，无需显式关闭文档类型
                System.Windows.MessageBox.Show("已导出图片");
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("导出失败：" + ex.Message); }
        }
    }

    private void BtnPickMerge_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = true };
        if (ofd.ShowDialog() == true)
        {
            _merges.Clear();
            _merges.AddRange(ofd.FileNames);
            MergeList.ItemsSource = null;
            MergeList.ItemsSource = _merges;
        }
    }

    private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
    {
        if (MergeList.SelectedItem is string sel)
        {
            int idx = _merges.IndexOf(sel);
            if (idx > 0)
            {
                _merges.RemoveAt(idx);
                _merges.Insert(idx - 1, sel);
                MergeList.ItemsSource = null;
                MergeList.ItemsSource = _merges;
                MergeList.SelectedIndex = idx - 1;
            }
        }
    }

    private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
    {
        if (MergeList.SelectedItem is string sel)
        {
            int idx = _merges.IndexOf(sel);
            if (idx >= 0 && idx < _merges.Count - 1)
            {
                _merges.RemoveAt(idx);
                _merges.Insert(idx + 1, sel);
                MergeList.ItemsSource = null;
                MergeList.ItemsSource = _merges;
                MergeList.SelectedIndex = idx + 1;
            }
        }
    }

    private void BtnDoMerge_Click(object sender, RoutedEventArgs e)
    {
        var files = _merges.ToArray();
        if (files.Length == 0) { System.Windows.MessageBox.Show("请先选择需要合并的 PDF"); return; }
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"合并_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            try { PdfMergeSplitService.Merge(files, sfd.FileName); System.Windows.MessageBox.Show("已合并"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("合并失败：" + ex.Message); }
        }
    }

    private void BtnPickSplit_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
        if (ofd.ShowDialog() == true)
        {
            _splitSrc = ofd.FileName;
        }
    }

    private void BtnDoSplit_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_splitSrc)) { System.Windows.MessageBox.Show("请先选择需要拆分的 PDF"); return; }
        if (!int.TryParse(SplitFrom.Text, out int from)) { System.Windows.MessageBox.Show("起始页无效"); return; }
        if (!int.TryParse(SplitTo.Text, out int to)) { System.Windows.MessageBox.Show("结束页无效"); return; }
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"拆分_{from}-{to}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            try { PdfMergeSplitService.Split(_splitSrc!, sfd.FileName, from, to); System.Windows.MessageBox.Show("已拆分"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("拆分失败：" + ex.Message); }
        }
    }

    private void BtnDoSplitEvery_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_splitSrc)) { System.Windows.MessageBox.Show("请先选择需要拆分的 PDF"); return; }
        if (ChkSplitEvery.IsChecked != true) { System.Windows.MessageBox.Show("请勾选‘每 N 页一份’"); return; }
        if (!int.TryParse(SplitEveryN.Text, out int n) || n <= 0) { System.Windows.MessageBox.Show("N 无效"); return; }
        var fbd = new System.Windows.Forms.FolderBrowserDialog();
        if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            try { PdfMergeSplitService.SplitEvery(_splitSrc!, fbd.SelectedPath, n); System.Windows.MessageBox.Show("已按 N 拆分"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("拆分失败：" + ex.Message); }
        }
    }

    private void BtnPickRotate_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
        if (ofd.ShowDialog() == true) _rotateSrc = ofd.FileName;
    }

    private void BtnDoRotate_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_rotateSrc)) { System.Windows.MessageBox.Show("请先选择 PDF"); return; }
        int angle = int.Parse(((System.Windows.Controls.ComboBoxItem)RotateAngleCombo.SelectedItem).Content!.ToString()!);
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"旋转_{angle}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            try { RotateService.Rotate(_rotateSrc!, sfd.FileName, angle); System.Windows.MessageBox.Show("已旋转"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("旋转失败：" + ex.Message); }
        }
    }

    private void BtnPickWatermark_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
        if (ofd.ShowDialog() == true) _watermarkSrc = ofd.FileName;
    }

    private void BtnDoWatermark_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_watermarkSrc)) { System.Windows.MessageBox.Show("请先选择 PDF"); return; }
        string text = WatermarkText.Text.Trim();
        if (string.IsNullOrEmpty(text)) { System.Windows.MessageBox.Show("请输入水印文本"); return; }
        bool diagonal = WatermarkDiagonal.IsChecked == true;
        double opacity = 0.15; double.TryParse(WatermarkOpacity.Text, out opacity);
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"水印_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            try { WatermarkService.AddTextWatermark(_watermarkSrc!, sfd.FileName, text, diagonal, opacity); System.Windows.MessageBox.Show("已添加水印"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("添加失败：" + ex.Message); }
        }
    }

    private void BtnPickPageNumber_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "PDF 文件|*.pdf", Multiselect = false };
        if (ofd.ShowDialog() == true) _pagenumSrc = ofd.FileName;
    }

    private void BtnDoPageNumber_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_pagenumSrc)) { System.Windows.MessageBox.Show("请先选择 PDF"); return; }
        string fmt = PageNumberFormat.Text.Trim(); if (string.IsNullOrEmpty(fmt)) fmt = "{page}/{total}";
        string posName = ((System.Windows.Controls.ComboBoxItem)PageNumberPos.SelectedItem).Content!.ToString()!;
        Enum.TryParse<PageNumberPosition>(posName, out var pos);
        var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "PDF 文件|*.pdf", FileName = $"页码_{DateTime.Now:yyyyMMdd_HHmmss}.pdf" };
        if (sfd.ShowDialog() == true)
        {
            try { PageNumberService.AddPageNumbers(_pagenumSrc!, sfd.FileName, fmt, pos); System.Windows.MessageBox.Show("已添加页码"); }
            catch (Exception ex) { System.Windows.MessageBox.Show("添加失败：" + ex.Message); }
        }
    }

    // 压缩相关事件处理已移除
}
