using System;
using System.IO;
using System.Text.Json;
using Pdftools.Core.Models;
using System.Threading;
using Pdftools.Core.Utilities;

namespace Pdftools.Core.Services;

/// <summary>
/// 应用设置读写服务：持久化默认打印机、默认模板等
/// </summary>
public static class SettingsService
{
    private static readonly Mutex _mtx = new(false, "Global/pdftools_settings_mutex");
    public static AppSettings Load()
    {
        try
        {
            Directory.CreateDirectory(Paths.AppDataDir);
            if (!File.Exists(Paths.ConfigFile))
                return new AppSettings();
            if (_mtx.WaitOne(TimeSpan.FromSeconds(3)))
            {
                try
                {
                    var json = File.ReadAllText(Paths.ConfigFile);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
                finally { try { _mtx.ReleaseMutex(); } catch { } }
            }
            else
            {
                // 超时回退
                var json = File.ReadAllText(Paths.ConfigFile);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
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
        var tmp = Paths.ConfigFile + ".tmp";
        if (_mtx.WaitOne(TimeSpan.FromSeconds(5)))
        {
            try
            {
                File.WriteAllText(tmp, json);
                // 原子替换（.NET 6+/Windows 支持覆盖移动）
                try { File.Move(tmp, Paths.ConfigFile, overwrite: true); }
                catch
                {
                    // 回退：复制覆盖
                    File.Copy(tmp, Paths.ConfigFile, overwrite: true);
                    try { File.Delete(tmp); } catch { }
                }
            }
            finally { try { _mtx.ReleaseMutex(); } catch { } }
        }
        else
        {
            // 超时回退：尽力而为
            File.WriteAllText(tmp, json);
            try { File.Move(tmp, Paths.ConfigFile, overwrite: true); } catch { File.Copy(tmp, Paths.ConfigFile, overwrite: true); try { File.Delete(tmp); } catch { } }
        }
    }
}
