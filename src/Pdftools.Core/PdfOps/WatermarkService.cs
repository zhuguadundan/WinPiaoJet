using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Pdftools.Core.PdfOps;

/// <summary>
/// 文本水印（简单版）：中心或斜向
/// </summary>
public static class WatermarkService
{
    /// <param name="text">水印文本</param>
    /// <param name="diagonal">是否斜向（true=沿对角线）</param>
    /// <param name="opacity">透明度 0..1</param>
    /// <param name="fontSize">字号（pt）</param>
    public static void AddTextWatermark(string input, string output, string text, bool diagonal = true, double opacity = 0.15, double fontSize = 64)
    {
        using var doc = PdfReader.Open(input, PdfDocumentOpenMode.Modify);
        for (int i = 0; i < doc.PageCount; i++)
        {
            var page = doc.Pages[i];
            using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
            double w = page.Width.Point, h = page.Height.Point;
            // 嵌入字体以提升跨机显示一致性
            var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.EmbedCompleteFontFile);
            var font = new XFont("Arial", fontSize, XFontStyleEx.Bold, fontOptions);
            var color = XColor.FromArgb((int)(opacity * 255), 200, 0, 0);
            var brush = new XSolidBrush(color);

            gfx.Save();
            if (diagonal)
            {
                gfx.TranslateTransform(w / 2, h / 2);
                gfx.RotateTransform(-45);
                var size = gfx.MeasureString(text, font);
                gfx.DrawString(text, font, brush, new XPoint(-size.Width / 2, size.Height / 2));
            }
            else
            {
                var size = gfx.MeasureString(text, font);
                gfx.DrawString(text, font, brush, new XPoint((w - size.Width) / 2, (h + size.Height) / 2));
            }
            gfx.Restore();
        }
        doc.Save(output);
    }
}
