using System;

namespace Pdftools.Core.Services;

/// <summary>
/// 打印布局与缩放计算：A5→A4 上半张，优先 1:1，不放大，仅在边距受限时 90%~100% 微缩。
/// 单位说明：
/// - 输入页面尺寸：pt（1pt=1/72in）
/// - 打印机可打印区域与边距：mm
/// - 结果：缩放系数（相对 1:1）、在可打印区域内的偏移（mm）
/// </summary>
public static class PrintLayout
{
    private const double PtToMm = 25.4 / 72.0; // 1pt -> mm

    public record Result(double Scale /*相对1:1*/, double OffsetLeftMm, double OffsetTopMm, double ContentWidthMm, double ContentHeightMm);

    /// <summary>
    /// 计算 A5 内容放置到 A4 上半张的定位与缩放（自动微缩启用）。
    /// </summary>
    public static Result ComputeA5Top(double pageWidthPt, double pageHeightPt,
        double printableWidthMm, double printableHeightMm,
        double marginMm, double safeGapMm)
    {
        return ComputeA5Top(pageWidthPt, pageHeightPt, printableWidthMm, printableHeightMm, marginMm, safeGapMm, true);
    }

    /// <summary>
    /// 计算 A5 内容放置到 A4 上半张的定位与缩放。
    /// </summary>
    /// <param name="pageWidthPt">PDF 页面宽（pt）</param>
    /// <param name="pageHeightPt">PDF 页面高（pt）</param>
    /// <param name="printableWidthMm">打印机可打印宽（mm）</param>
    /// <param name="printableHeightMm">打印机可打印高（mm）</param>
    /// <param name="marginMm">边距预设（mm），应用于上半区域的内边距</param>
    /// <param name="safeGapMm">中线安全距（mm），防止溢出下半张</param>
    /// <param name="autoShrink">是否启用 90%~100% 自动微缩</param>
    /// <returns>缩放与偏移</returns>
    public static Result ComputeA5Top(double pageWidthPt, double pageHeightPt,
        double printableWidthMm, double printableHeightMm,
        double marginMm, double safeGapMm, bool autoShrink)
    {
        var pageMmW = pageWidthPt * PtToMm;
        var pageMmH = pageHeightPt * PtToMm;

        // 目标区域（上半张，锚点左上）
        double targetWidthMm = Math.Max(0, printableWidthMm - marginMm * 2);
        double targetHeightMm = Math.Max(0, printableHeightMm / 2.0 - safeGapMm - marginMm);

        // 1:1 不放大
        double scale = 1.0;
        bool fits1 = pageMmW <= targetWidthMm && pageMmH <= targetHeightMm;
        if (!fits1 && autoShrink)
        {
            // 等比：取 min(宽比, 高比)，但不超过 1.00，不低于 0.90
            double sw = targetWidthMm / pageMmW;
            double sh = targetHeightMm / pageMmH;
            double s = Math.Min(sw, sh);
            s = Math.Min(1.0, Math.Max(0.90, s));
            scale = s;
        }

        double contentW = pageMmW * scale;
        double contentH = pageMmH * scale;

        // 左上锚点：偏移=边距
        double offsetLeft = marginMm;
        double offsetTop = marginMm; // 相对于可打印区域的上边

        return new Result(scale, offsetLeft, offsetTop, contentW, contentH);
    }
}
