using System;

namespace Pdftools.Core.Services;

/// <summary>
/// 占位实现：后续用 PDFiumCore 实现真实渲染与尺寸识别
/// </summary>
public class PdfRendererStub : IPdfRenderer
{
    public (double widthPt, double heightPt) GetPageSize(string filePath, int pageIndex = 0)
    {
        throw new NotImplementedException("尚未实现：使用 PDFiumCore 获取页面尺寸");
    }

    public bool IsA5Like(string filePath, int pageIndex = 0, double tolerance = 0.05)
    {
        // 占位：直接返回 true，防止调用方阻塞
        return true;
    }
}
