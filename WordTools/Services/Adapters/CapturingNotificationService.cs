using System.Collections.Generic;
using System.Windows.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    public sealed class CapturedNotification
    {
        public string Message { get; set; }
        public string Title { get; set; }
    }

    /// <summary>
    /// 捕获 MessageBox 文案的通知服务，供自动化断言；默认确认对话框返回 Yes。
    /// </summary>
    public sealed class CapturingNotificationService : INotificationService
    {
        public List<CapturedNotification> Warnings { get; } = new List<CapturedNotification>();
        public List<CapturedNotification> Informations { get; } = new List<CapturedNotification>();
        public List<CapturedNotification> Errors { get; } = new List<CapturedNotification>();
        public List<CapturedNotification> Questions { get; } = new List<CapturedNotification>();
        public DialogResult QuestionResult { get; set; } = DialogResult.Yes;

        public void ShowInformation(string message, string title)
        {
            Informations.Add(new CapturedNotification { Message = message, Title = title });
        }

        public void ShowWarning(string message, string title)
        {
            Warnings.Add(new CapturedNotification { Message = message, Title = title });
        }

        public void ShowError(string message, string title)
        {
            Errors.Add(new CapturedNotification { Message = message, Title = title });
        }

        public DialogResult ShowQuestion(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            Questions.Add(new CapturedNotification { Message = message, Title = title });
            return QuestionResult;
        }
    }
}
