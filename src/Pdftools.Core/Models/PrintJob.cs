namespace Pdftools.Core.Models;

/// <summary>
/// 打印任务模型（精简版占位）
/// </summary>
public class PrintJob
{
    /// <summary>源文件路径</summary>
    public string FilePath { get; set; } = string.Empty;
    /// <summary>页范围，如 "all" 或 "1-3,5"</summary>
    public string PageRange { get; set; } = "all";
    /// <summary>模板 ID</summary>
    public string TemplateId { get; set; } = "default";
    /// <summary>状态：Pending/Running/Done/Failed</summary>
    public string Status { get; set; } = "Pending";
    /// <summary>错误消息</summary>
    public string? ErrorMessage { get; set; }
}
