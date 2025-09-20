using System;
using System.IO;
using System.Text.Json;
using Pdftools.Core.Models;
using Pdftools.Core.Utilities;

namespace Pdftools.Core.Services;

/// <summary>
/// 应用设置读写服务：持久化默认打印机、默认模板等
/// </summary>
public static class SettingsService
{
    public static AppSettings Load()
    {
        try
        {
            Directory.CreateDirectory(Paths.AppDataDir);
            if (!File.Exists(Paths.ConfigFile))
                return new AppSettings();
            var json = File.ReadAllText(Paths.ConfigFile);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        Directory.CreateDirectory(Paths.AppDataDir);
        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(Paths.ConfigFile, json);
    }
}
