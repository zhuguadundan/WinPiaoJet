using System;
using System.Diagnostics;
using System.IO;

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
            p.WaitForExit(2000);
            return p.ExitCode == 0 || p.ExitCode == 2; // qpdf 有时返回 2 但仍输出版本
        }
        catch { return false; }
    }

    public static void Compress(string input, string output, bool linearize = true)
    {
        if (!File.Exists(input)) throw new FileNotFoundException("输入文件不存在", input);
        if (!IsAvailable()) throw new InvalidOperationException("未检测到 qpdf，请先安装并加入 PATH。");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
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
}
