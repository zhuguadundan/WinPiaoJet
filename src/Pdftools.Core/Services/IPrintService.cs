namespace Pdftools.Core.Services;

/// <summary>
/// 打印服务接口（基于 System.Printing 封装）
/// </summary>
public interface IPrintService
{
    /// <summary>获取可用打印机列表</summary>
    string[] GetPrinters();
    /// <summary>打印单个 PDF 文件至 A4 上半张（1:1 优先 + 自动微缩）</summary>
    void PrintA5ToA4Top(string filePath, string printerName, string templateId, int dpi = 300);
}
