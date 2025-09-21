using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Pdftools.Desktop.Services;

/// <summary>
/// QPDF 压缩与线性化调用器（依赖本机已安装 qpdf.exe）
/// </summary>
public static class QpdfRunner
{
    public static bool IsAvailable()
    {
        try
        {
            var psi = new ProcessStartInfo("qpdf", "--version") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            using var p = Process.Start(psi)!;
            var exited = p.WaitForExit(2000);
            if (!exited)
            {
                try { p.Kill(entireProcessTree: true); } catch { }
                return false;
            }
            var code = 0;
            try { code = p.ExitCode; } catch { return false; }
            return code == 0 || code == 2; // qpdf 有时返回 2 但仍输出版本
        }
        catch { return false; }
    }

    public static void Compress(string input, string output, bool linearize = true)
    {
        if (!File.Exists(input)) throw new FileNotFoundException("输入文件不存在", input);
        if (!IsAvailable()) throw new InvalidOperationException("未检测到 qpdf，请先安装并加入 PATH。");
        var outDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);
        var args = linearize ? $"--linearize --object-streams=generate --recompress-flate \"{input}\" \"{output}\""
                             : $"--object-streams=generate --recompress-flate \"{input}\" \"{output}\"";
        var psi = new ProcessStartInfo("qpdf", args)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi)!;
        p.WaitForExit();
        if (p.ExitCode != 0)
        {
            string err = p.StandardError.ReadToEnd();
            throw new Exception("qpdf 执行失败：" + err);
        }
    }

    /// <summary>
    /// 异步压缩，支持可选超时（默认 2 分钟）。
    /// </summary>
    public static async Task CompressAsync(string input, string output, bool linearize = true, TimeSpan? timeout = null, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(input)) throw new FileNotFoundException("输入文件不存在", input);
        if (!IsAvailable()) throw new InvalidOperationException("未检测到 qpdf，请先安装并加入 PATH。");
        var outDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);

        var args = linearize ? $"--linearize --object-streams=generate --recompress-flate \"{input}\" \"{output}\""
                             : $"--object-streams=generate --recompress-flate \"{input}\" \"{output}\"";
        var psi = new ProcessStartInfo("qpdf", args)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi)!;
        var to = timeout ?? TimeSpan.FromMinutes(2);
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(to);
        try
        {
            await p.WaitForExitAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            try { p.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"qpdf 执行超时（{to}）");
        }
        if (p.ExitCode != 0)
        {
            string err = await p.StandardError.ReadToEndAsync();
            throw new Exception("qpdf 执行失败：" + err);
        }
    }
}
