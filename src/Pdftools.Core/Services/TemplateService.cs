using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Pdftools.Core.Models;
using Pdftools.Core.Utilities;

namespace Pdftools.Core.Services;

/// <summary>
/// 模板管理（文件持久化，简化实现）
/// </summary>
public class TemplateService : ITemplateService
{
    private readonly string _filePath;
    private List<Template> _templates;

    public TemplateService()
    {
        Directory.CreateDirectory(Paths.AppDataDir);
        _filePath = Path.Combine(Paths.AppDataDir, "templates.json");
        _templates = Load();
        if (_templates.Count == 0)
        {
            _templates.Add(new Template { Id = "default", Name = "默认模板", MarginMm = 3, AutoShrink = true, SafeGapMm = 5 });
            SaveAll();
        }
    }

    public Template Get(string id) => _templates.FirstOrDefault(t => t.Id == id) ?? _templates[0];

    public Template[] GetAll() => _templates.ToArray();

    public void Save(Template template)
    {
        var existing = _templates.FindIndex(t => t.Id == template.Id);
        if (existing >= 0) _templates[existing] = template; else _templates.Add(template);
        SaveAll();
    }

    public void Delete(string id)
    {
        _templates.RemoveAll(t => t.Id == id && t.Id != "default");
        SaveAll();
    }

    private List<Template> Load()
    {
        if (!File.Exists(_filePath)) return new List<Template>();
        try
        {
            var json = File.ReadAllText(_filePath);
            return JsonSerializer.Deserialize<List<Template>>(json) ?? new List<Template>();
        }
        catch
        {
            return new List<Template>();
        }
    }

    private void SaveAll()
    {
        var json = JsonSerializer.Serialize(_templates, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_filePath, json);
    }
}
