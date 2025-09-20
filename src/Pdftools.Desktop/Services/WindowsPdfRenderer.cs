using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

namespace Pdftools.Desktop.Services;

public static class WindowsPdfRenderer
{
    public static async Task<BitmapSource> RenderPageAsync(string filePath, int pageIndex, int dpi)
    {
        StorageFile sf = await StorageFile.GetFileFromPathAsync(filePath);
        var doc = await PdfDocument.LoadFromFileAsync(sf);
        if (doc == null || doc.PageCount == 0) throw new InvalidOperationException("PDF 无效");
        var page = doc.GetPage((uint)Math.Min(pageIndex, (int)doc.PageCount - 1));

        using var stream = new InMemoryRandomAccessStream();
        var dims = page.Dimensions;
        // Windows.Data.Pdf 页面尺寸以 1/96" 为单位（DIP），按 dpi 比例缩放
        double widthDip = dims.MediaBox.Width;
        double heightDip = dims.MediaBox.Height;
        var options = new PdfPageRenderOptions
        {
            DestinationWidth = (uint)Math.Max(1, Math.Round(widthDip * dpi / 96.0)),
            DestinationHeight = (uint)Math.Max(1, Math.Round(heightDip * dpi / 96.0))
        };
        await page.RenderToStreamAsync(stream, options);

        // 转 BitmapImage
        stream.Seek(0);
        using var netStream = stream.AsStream();
        var ms = new MemoryStream();
        await netStream.CopyToAsync(ms);
        ms.Position = 0;
        var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
        return decoder.Frames[0];
    }
}
