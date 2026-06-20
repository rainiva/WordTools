using System.Windows.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// INotificationService 的 MessageBox 实现。
    /// </summary>
    public sealed class MessageBoxNotificationService : INotificationService
    {
        public void ShowInformation(string message, string title)
            => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        public void ShowWarning(string message, string title)
            => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        public void ShowError(string message, string title)
            => MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        public DialogResult ShowQuestion(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
            => MessageBox.Show(message, title, buttons, icon);
    }
}
