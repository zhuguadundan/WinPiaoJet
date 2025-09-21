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
        using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return await RenderPageAsync(fs, pageIndex, dpi);
    }

    public static async Task<BitmapSource> RenderPageAsync(string filePath, int pageIndex, int pixelWidth, int pixelHeight)
    {
        using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return await RenderPageAsync(fs, pageIndex, pixelWidth, pixelHeight);
    }

    /// <summary>
    /// 从任意流渲染指定页到位图（按 DPI）
    /// </summary>
    public static async Task<BitmapSource> RenderPageAsync(Stream pdfStream, int pageIndex, int dpi)
    {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        if (!pdfStream.CanRead) throw new ArgumentException("不可读取的流", nameof(pdfStream));

        // 将托管流复制到 WinRT 内存流
        using var ra = new InMemoryRandomAccessStream();
        if (pdfStream.CanSeek) pdfStream.Position = 0;
        using (var bufferMs = new MemoryStream())
        {
            await pdfStream.CopyToAsync(bufferMs).ConfigureAwait(false);
            var writer = new DataWriter(ra);
            var bytes = bufferMs.ToArray();
            writer.WriteBytes(bytes);
            await writer.StoreAsync();
            writer.DetachStream();
            writer.Dispose();
        }
        ra.Seek(0);

        // 加载文档与页面、渲染到 PNG 内存流
        var doc = await PdfDocument.LoadFromStreamAsync(ra);
        if (doc == null || doc.PageCount == 0) throw new InvalidOperationException("PDF 无效");
        using (var page = doc.GetPage((uint)Math.Min(pageIndex, (int)doc.PageCount - 1)))
        using (var stream = new InMemoryRandomAccessStream())
        {
            var dims = page.Dimensions; // DIP (1/96")
            double widthDip = dims.MediaBox.Width;
            double heightDip = dims.MediaBox.Height;
            var options = new PdfPageRenderOptions
            {
                DestinationWidth = (uint)Math.Max(1, Math.Round(widthDip * dpi / 96.0)),
                DestinationHeight = (uint)Math.Max(1, Math.Round(heightDip * dpi / 96.0))
            };
            await page.RenderToStreamAsync(stream, options);
            return await DecodeToBitmapAsync(stream);
        }
    }

    /// <summary>
    /// 从任意流渲染指定页到位图（按目标像素尺寸）
    /// </summary>
    public static async Task<BitmapSource> RenderPageAsync(Stream pdfStream, int pageIndex, int pixelWidth, int pixelHeight)
    {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        if (!pdfStream.CanRead) throw new ArgumentException("不可读取的流", nameof(pdfStream));

        using var ra = new InMemoryRandomAccessStream();
        if (pdfStream.CanSeek) pdfStream.Position = 0;
        using (var bufferMs = new MemoryStream())
        {
            await pdfStream.CopyToAsync(bufferMs).ConfigureAwait(false);
            var writer = new DataWriter(ra);
            var bytes = bufferMs.ToArray();
            writer.WriteBytes(bytes);
            await writer.StoreAsync();
            writer.DetachStream();
            writer.Dispose();
        }
        ra.Seek(0);

        var doc = await PdfDocument.LoadFromStreamAsync(ra);
        if (doc == null || doc.PageCount == 0) throw new InvalidOperationException("PDF 无效");
        using (var page = doc.GetPage((uint)Math.Min(pageIndex, (int)doc.PageCount - 1)))
        using (var stream = new InMemoryRandomAccessStream())
        {
            var options = new PdfPageRenderOptions
            {
                DestinationWidth = (uint)Math.Max(1, pixelWidth),
                DestinationHeight = (uint)Math.Max(1, pixelHeight)
            };
            await page.RenderToStreamAsync(stream, options);
            return await DecodeToBitmapAsync(stream);
        }
    }
    // 移除返回 doc/page 的占位实现，避免文档与流泄漏；渲染方法内部自管理生命周期

    public static async Task<(byte[] data, int width, int height)> RenderPagePngAsync(string filePath, int pageIndex, int pixelWidth, int pixelHeight)
    {
        using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return await RenderPagePngAsync(fs, pageIndex, pixelWidth, pixelHeight);
    }

    public static async Task<(byte[] data, int width, int height)> RenderPagePngAsync(Stream pdfStream, int pageIndex, int pixelWidth, int pixelHeight)
    {
        using var ra = new InMemoryRandomAccessStream();
        if (pdfStream.CanSeek) pdfStream.Position = 0;
        using (var bufferMs = new MemoryStream())
        {
            await pdfStream.CopyToAsync(bufferMs).ConfigureAwait(false);
            var writer = new DataWriter(ra);
            var bytes = bufferMs.ToArray();
            writer.WriteBytes(bytes);
            await writer.StoreAsync();
            writer.DetachStream();
            writer.Dispose();
        }
        ra.Seek(0);
        var doc = await PdfDocument.LoadFromStreamAsync(ra);
        if (doc == null || doc.PageCount == 0) throw new InvalidOperationException("PDF 无效");
        using (var page = doc.GetPage((uint)Math.Min(pageIndex, (int)doc.PageCount - 1)))
        using (var stream = new InMemoryRandomAccessStream())
        {
            var options = new PdfPageRenderOptions
            {
                DestinationWidth = (uint)Math.Max(1, pixelWidth),
                DestinationHeight = (uint)Math.Max(1, pixelHeight)
            };
            await page.RenderToStreamAsync(stream, options);
            // 读出字节，不解码（避免线程亲缘问题）
            stream.Seek(0);
            var reader = new DataReader(stream.GetInputStreamAt(0));
            uint size = (uint)stream.Size;
            await reader.LoadAsync(size);
            var outBytes = new byte[size];
            reader.ReadBytes(outBytes);
            reader.Dispose();
            return (outBytes, pixelWidth, pixelHeight);
        }
    }

    private static async Task<BitmapSource> DecodeToBitmapAsync(InMemoryRandomAccessStream stream)
    {
        stream.Seek(0);
        var reader = new DataReader(stream.GetInputStreamAt(0));
        uint size = (uint)stream.Size;
        await reader.LoadAsync(size);
        var bytes = new byte[size];
        reader.ReadBytes(bytes);
        reader.Dispose();
        using var ms = new MemoryStream(bytes);
        var decoder = BitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
        var frame = decoder.Frames[0];
        return frame; // 不在此处 Freeze，留给调用方在所属线程处理
    }
}
