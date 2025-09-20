using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Pdftools.Core.PdfOps;

/// <summary>
/// PDF 合并与拆分（基础版）
/// </summary>
public static class PdfMergeSplitService
{
    public static void Merge(string[] inputs, string output)
    {
        using var outDoc = new PdfDocument();
        foreach (var file in inputs)
        {
            using var src = PdfReader.Open(file, PdfDocumentOpenMode.Import);
            for (int i = 0; i < src.PageCount; i++)
            {
                outDoc.AddPage(src.Pages[i]);
            }
        }
        outDoc.Save(output);
    }

    /// <summary>
    /// 拆分指定页区间（起止为 1-based，包含）
    /// </summary>
    public static void Split(string input, string output, int fromPage, int toPage)
    {
        using var src = PdfReader.Open(input, PdfDocumentOpenMode.Import);
        using var outDoc = new PdfDocument();
        fromPage = Math.Max(1, fromPage);
        toPage = Math.Min(src.PageCount, toPage);
        for (int i = fromPage - 1; i < toPage; i++)
        {
            outDoc.AddPage(src.Pages[i]);
        }
        outDoc.Save(output);
    }

    /// <summary>
    /// 按每 N 页拆分为多份，输出到文件夹，文件名模式例如 "split_{index}.pdf"
    /// </summary>
    public static void SplitEvery(string input, string outFolder, int pagesPer, string fileNamePattern = "split_{index}.pdf")
    {
        if (pagesPer <= 0) throw new ArgumentException("pagesPer 必须 > 0");
        Directory.CreateDirectory(outFolder);
        using var src = PdfReader.Open(input, PdfDocumentOpenMode.Import);
        int total = src.PageCount;
        int index = 1;
        for (int start = 0; start < total; start += pagesPer)
        {
            using var outDoc = new PdfDocument();
            int end = Math.Min(total, start + pagesPer);
            for (int i = start; i < end; i++)
                outDoc.AddPage(src.Pages[i]);
            string name = fileNamePattern.Replace("{index}", index.ToString());
            string path = Path.Combine(outFolder, name);
            outDoc.Save(path);
            index++;
        }
    }
}
