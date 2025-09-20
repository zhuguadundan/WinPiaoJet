using System.Linq;
using System.Windows;
using Pdftools.Core.Services;

namespace Pdftools.Desktop.Windows.Templates;

public partial class TemplatesWindow : Window
{
    private readonly TemplateService _svc = new();
    public TemplatesWindow()
    {
        InitializeComponent();
        LoadList();
    }

    private void LoadList()
    {
        TemplateList.ItemsSource = _svc.GetAll();
    }

    private void BtnRename_Click(object sender, RoutedEventArgs e)
    {
        if (TemplateList.SelectedItem is Pdftools.Core.Models.Template tpl)
        {
            var name = RenameBox.Text.Trim();
            if (string.IsNullOrEmpty(name)) { System.Windows.MessageBox.Show("请输入新名称"); return; }
            tpl.Name = name;
            _svc.Save(tpl);
            LoadList();
        }
    }

    private void BtnSetDefault_Click(object sender, RoutedEventArgs e)
    {
        if (TemplateList.SelectedItem is Pdftools.Core.Models.Template tpl)
        {
            var settings = SettingsService.Load();
            settings.DefaultTemplateId = tpl.Id;
            SettingsService.Save(settings);
            System.Windows.MessageBox.Show("已设为默认模板");
        }
    }

    private void BtnDelete_Click(object sender, RoutedEventArgs e)
    {
        if (TemplateList.SelectedItem is Pdftools.Core.Models.Template tpl)
        {
            if (tpl.Id == "default") { System.Windows.MessageBox.Show("默认模板不可删除"); return; }
            _svc.Delete(tpl.Id);
            LoadList();
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
