using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Pdftools.Core.PdfOps;

public enum PageNumberPosition
{
    TopLeft, TopCenter, TopRight,
    BottomLeft, BottomCenter, BottomRight
}

/// <summary>
/// 为 PDF 添加页码（简单版）
/// </summary>
public static class PageNumberService
{
    /// <param name="format">例如 "{page}/{total}"、"第 {page} 页"</param>
    /// <param name="pos">位置九宫格之一</param>
    /// <param name="marginMm">边距（mm）</param>
    /// <param name="startAt">起始页码，默认 1</param>
    public static void AddPageNumbers(string input, string output, string format = "{page}/{total}", PageNumberPosition pos = PageNumberPosition.BottomCenter, int marginMm = 10, int startAt = 1)
    {
        using var doc = PdfReader.Open(input, PdfDocumentOpenMode.Modify);
        int total = doc.PageCount;
        double mmToPt = 72.0 / 25.4;
        double marginPt = marginMm * mmToPt;
        // 嵌入字体以提升跨机显示一致性
        var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.EmbedCompleteFontFile);
        var font = new XFont("Arial", 10, XFontStyleEx.Regular, fontOptions);
        var brush = XBrushes.Black;

        for (int i = 0; i < total; i++)
        {
            var page = doc.Pages[i];
            using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            string text = format.Replace("{page}", (startAt + i).ToString()).Replace("{total}", total.ToString());
            var size = gfx.MeasureString(text, font);
            double x = marginPt, y = marginPt;
            double w = page.Width.Point, h = page.Height.Point;
            switch (pos)
            {
                case PageNumberPosition.TopLeft:
                    x = marginPt; y = marginPt + size.Height; break;
                case PageNumberPosition.TopCenter:
                    x = (w - size.Width) / 2; y = marginPt + size.Height; break;
                case PageNumberPosition.TopRight:
                    x = w - marginPt - size.Width; y = marginPt + size.Height; break;
                case PageNumberPosition.BottomLeft:
                    x = marginPt; y = h - marginPt; break;
                case PageNumberPosition.BottomCenter:
                    x = (w - size.Width) / 2; y = h - marginPt; break;
                case PageNumberPosition.BottomRight:
                    x = w - marginPt - size.Width; y = h - marginPt; break;
            }
            gfx.DrawString(text, font, brush, new XPoint(x, y));
        }
        doc.Save(output);
    }
}
