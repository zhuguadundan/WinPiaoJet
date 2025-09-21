using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;

namespace Pdftools.Tester;

internal class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string root = AppContext.BaseDirectory;
            // 仓库根目录定位
            var dir = new DirectoryInfo(root);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "Pdftools.sln"))) dir = dir.Parent;
            string repo = dir?.FullName ?? Environment.CurrentDirectory;

            string? file = args.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(file))
            {
                file = Directory.EnumerateFiles(repo, "*.pdf", SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
            {
                Console.WriteLine("未找到 PDF 测试文件（请在仓库根目录放入发票 PDF）");
                return 2;
            }

            Console.WriteLine($"Testing file: {file}");

            // 1) 渲染第一页为位图
            var bs = Pdftools.Desktop.Services.WindowsPdfRenderer.RenderPageAsync(file, 0, 300)
                .GetAwaiter().GetResult();
            if (bs.CanFreeze) bs.Freeze();

            // 保存 PNG 以验证渲染（通过 RenderTargetBitmap 复制到当前线程拥有的位图）
            string outDir = Path.Combine(repo, "dist");
            Directory.CreateDirectory(outDir);
            string png = Path.Combine(outDir, "test-preview.png");
            // 先复制到当前线程拥有的 RenderTargetBitmap
            var dvCopy = new DrawingVisual();
            using (var dc = dvCopy.RenderOpen())
            {
                dc.DrawImage(bs, new System.Windows.Rect(0, 0, bs.PixelWidth, bs.PixelHeight));
            }
            var copyRtb = new RenderTargetBitmap(bs.PixelWidth, bs.PixelHeight, 96, 96, PixelFormats.Pbgra32);
            copyRtb.Render(dvCopy);

            using (var fs = File.Open(png, FileMode.Create, FileAccess.Write))
            {
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(copyRtb));
                enc.Save(fs);
            }
            Console.WriteLine($"PNG saved: {png}");

            // 2) 模拟打印：写出 XPS 文件（用当前线程拥有的 rtb 作为来源）
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                // 简单写满 600x800 区域
                dc.DrawImage(copyRtb, new System.Windows.Rect(0, 0, 600, 800));
            }
            string xps = Path.Combine(outDir, "test-print.xps");
            if (File.Exists(xps)) File.Delete(xps);
            using (var xpsDoc = new XpsDocument(xps, FileAccess.ReadWrite))
            {
                var writer = XpsDocument.CreateXpsDocumentWriter(xpsDoc);
                writer.Write(visual);
            }
            Console.WriteLine($"XPS saved: {xps}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }
}
