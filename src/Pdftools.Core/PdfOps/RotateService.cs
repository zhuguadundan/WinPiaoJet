using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Pdftools.Core.PdfOps;

/// <summary>
/// 旋转页面（90/180/270）
/// </summary>
public static class RotateService
{
    public static void Rotate(string input, string output, int angle)
    {
        if (angle % 90 != 0) throw new System.ArgumentException("角度必须是 90 的倍数");
        using var doc = PdfReader.Open(input, PdfDocumentOpenMode.Modify);
        for (int i = 0; i < doc.PageCount; i++)
        {
            var page = doc.Pages[i];
            int current = page.Rotate;
            page.Rotate = (current + angle) % 360;
        }
        doc.Save(output);
    }
}
