using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Pdftools.Desktop.Models;

/// <summary>
/// 文件列表项（用于 DataGrid 绑定）
/// </summary>
public class FileItem : INotifyPropertyChanged
{
    private string _name = string.Empty;
    private int _pageCount;
    private string _size = string.Empty;
    private string _status = "待处理";
    private string _message = string.Empty;
    private string _path = string.Empty;

    public string Name { get => _name; set { _name = value; OnPropertyChanged(); } }
    public int PageCount { get => _pageCount; set { _pageCount = value; OnPropertyChanged(); } }
    public string Size { get => _size; set { _size = value; OnPropertyChanged(); } }
    public string Status { get => _status; set { _status = value; OnPropertyChanged(); } }
    public string Message { get => _message; set { _message = value; OnPropertyChanged(); } }
    public string Path { get => _path; set { _path = value; OnPropertyChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
