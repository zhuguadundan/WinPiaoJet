namespace Pdftools.Core.Models;

/// <summary>
/// 打印模板配置（与 PRD 对齐，简化字段）
/// </summary>
public class Template
{
    public string Id { get; set; } = "default";
    public string Name { get; set; } = "默认模板";
    /// <summary>边距预设（毫米）：0/3/5/8</summary>
    public int MarginMm { get; set; } = 3;
    /// <summary>是否启用自动微缩（90%~100%）</summary>
    public bool AutoShrink { get; set; } = true;
    /// <summary>安全距（毫米），默认 5</summary>
    public int SafeGapMm { get; set; } = 5;
}
