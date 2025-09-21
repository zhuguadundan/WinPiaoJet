using System;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;

namespace Pdftools.Desktop.Services
{
    /// <summary>
    /// 基于 Windows.Data.Pdf 的 IPdfRenderer 实现（最小可用）。
    /// 说明：PdfPage.Dimensions 的单位按现有注释视为 DIP(1/96")，换算为 pt 需乘以 72/96。
    /// </summary>
    public class PdfRendererWinRT : Pdftools.Core.Services.IPdfRenderer
    {
        private const double DipToPt = 72.0 / 96.0;
        private const double PtToMm = 25.4 / 72.0;

        public (double widthPt, double heightPt) GetPageSize(string filePath, int pageIndex = 0)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            StorageFile sf = StorageFile.GetFileFromPathAsync(filePath).AsTask().GetAwaiter().GetResult();
            var doc = PdfDocument.LoadFromFileAsync(sf).AsTask().GetAwaiter().GetResult();
            if (doc == null || doc.PageCount == 0) throw new InvalidOperationException("PDF 无效");
            using var page = doc.GetPage((uint)Math.Min(pageIndex, (int)doc.PageCount - 1));
            var dims = page.Dimensions;
            // 将 DIP 换算为 pt
            double wPt = dims.MediaBox.Width * DipToPt;
            double hPt = dims.MediaBox.Height * DipToPt;
            return (wPt, hPt);
        }

        public bool IsA5Like(string filePath, int pageIndex = 0, double tolerance = 0.05)
        {
            var (wPt, hPt) = GetPageSize(filePath, pageIndex);
            // 转 mm
            double wMm = wPt * PtToMm;
            double hMm = hPt * PtToMm;
            // A5: 148x210 mm（容许旋转）
            double a5w = 148.0, a5h = 210.0;
            bool matchNormal = Nearly(wMm, a5w, tolerance) && Nearly(hMm, a5h, tolerance);
            bool matchRot = Nearly(wMm, a5h, tolerance) && Nearly(hMm, a5w, tolerance);
            return matchNormal || matchRot;
        }

        private static bool Nearly(double v, double target, double tol)
        {
            // 相对误差容限
            if (target == 0) return Math.Abs(v) < 1e-6;
            return Math.Abs(v - target) / target <= tol;
        }
    }
}
