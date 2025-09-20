using System;
using System.IO;

namespace Pdftools.Core.Utilities;

/// <summary>
/// 路径助手：配置与日志路径
/// </summary>
public static class Paths
{
    public static string AppDataDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "pdftools");

    public static string LocalDataDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "pdftools");

    public static string ConfigFile => Path.Combine(AppDataDir, "config.json");

    public static string LogsDir => Path.Combine(LocalDataDir, "logs");
}
