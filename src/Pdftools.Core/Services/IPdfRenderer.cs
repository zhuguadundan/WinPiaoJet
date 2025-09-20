namespace Pdftools.Core.Services;

/// <summary>
/// PDF 渲染与尺寸识别服务接口（后续用 PDFium 实现）
/// </summary>
public interface IPdfRenderer
{
    /// <summary>获取第一页尺寸（点，1 点=1/72 英寸）</summary>
    (double widthPt, double heightPt) GetPageSize(string filePath, int pageIndex = 0);
    /// <summary>是否为 A5 或准 A5（±5%）</summary>
    bool IsA5Like(string filePath, int pageIndex = 0, double tolerance = 0.05);
}
