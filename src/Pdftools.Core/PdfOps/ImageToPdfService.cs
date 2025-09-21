using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Pdftools.Core.PdfOps;

/// <summary>
/// 图片→PDF（简单版）：
/// - 将多张 JPG/PNG/BMP 按顺序放入 PDF，每张图一页
/// - 页面：A4 纵向；等比缩放铺满页内（保留边距）
/// </summary>
public static class ImageToPdfService
{
    /// <param name="images">图片路径列表</param>
    /// <param name="outputPdf">输出 PDF 路径</param>
    /// <param name="marginMm">页边距（毫米）</param>
    public static void Convert(string[] images, string outputPdf, double marginMm = 10, PdfSharp.PageSize pageSize = PdfSharp.PageSize.A4)
    {
        using var doc = new PdfDocument();
        foreach (var imgPath in images)
        {
            var page = doc.AddPage();
            page.Size = pageSize;
            using var gfx = XGraphics.FromPdfPage(page);
            using var ximg = XImage.FromFile(imgPath);

            // 页面与边距（point，1 point=1/72 inch，1 inch=25.4mm）
            double mmToPt = 72.0 / 25.4;
            double marginPt = marginMm * mmToPt;
            double targetW = page.Width.Point - 2 * marginPt;
            double targetH = page.Height.Point - 2 * marginPt;

            // 等比缩放适配（使用图像的点单位，避免像素与 pt 混用）
            double imgWpt = ximg.PointWidth;
            double imgHpt = ximg.PointHeight;
            double scale = Math.Min(targetW / imgWpt, targetH / imgHpt);
            double drawW = imgWpt * scale;
            double drawH = imgHpt * scale;
            double x = marginPt + (targetW - drawW) / 2;
            double y = marginPt + (targetH - drawH) / 2;
            gfx.DrawImage(ximg, x, y, drawW, drawH);
        }
        doc.Save(outputPdf);
    }
}
