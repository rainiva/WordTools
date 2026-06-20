using System;
using Word = Microsoft.Office.Interop.Word;
using WordTools.Forms;
using WordTools.Services.Abstractions;
using WordTools.Services.Adapters;

namespace WordTools.Services
{
    /// <summary>
    /// Excel 数据填充流程编排：创建上下文并显示 ExcelDataFillerForm。
    /// </summary>
    public sealed class ExcelDataFillerOrchestrator
    {
        private readonly Word.Application _application;

        public ExcelDataFillerOrchestrator(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void ShowForm()
        {
            var notificationService = new MessageBoxNotificationService();

            try
            {
                var appContext = new WordApplicationContext(_application);
                var documentContext = new WordDocumentContext(appContext);

                using (var form = new ExcelDataFillerForm(documentContext, notificationService))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                notificationService.ShowError(
                    string.Format("打开Excel数据填充工具失败: {0}", ex.Message),
                    "错误");
            }
        }
    }
}
