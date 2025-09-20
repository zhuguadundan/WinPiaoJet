namespace Pdftools.Core.Services;

using Pdftools.Core.Models;

/// <summary>
/// 模板与预设管理接口
/// </summary>
public interface ITemplateService
{
    Template Get(string id);
    void Save(Template template);
    Template[] GetAll();
    void Delete(string id);
}
