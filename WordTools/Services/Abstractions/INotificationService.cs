using System.Windows.Forms;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 消息通知抽象。替代服务层直接调用 MessageBox。
    /// </summary>
    public interface INotificationService
    {
        void ShowInformation(string message, string title);
        void ShowWarning(string message, string title);
        void ShowError(string message, string title);
        DialogResult ShowQuestion(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon);
    }
}
