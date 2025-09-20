namespace Pdftools.Core.Models;

/// <summary>
/// 应用设置（持久化至 %AppData%\pdftools\config.json）
/// </summary>
public class AppSettings
{
    public string? DefaultPrinter { get; set; }
    public string[] FavoritePrinters { get; set; } = []; // 收藏打印机
    public string DefaultTemplateId { get; set; } = "default";
    public int Concurrency { get; set; } = 1; // 打印并发，默认 1
    public string LogLevel { get; set; } = "Information";
}
