namespace Pdftools.Core.Services;

using Pdftools.Core.Models;

/// <summary>
/// 打印任务队列接口（顺序处理，可暂停/继续/取消）
/// </summary>
public interface IJobQueue
{
    void Enqueue(PrintJob job);
    void Pause();
    void Resume();
    void CancelAll();
}
