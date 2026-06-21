using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordTools.Forms;
using WordTools.Services.Abstractions;
using WordTools.Services.Adapters;

namespace WordTools.Services
{
    /// <summary>
    /// 批量插图流程编排：窗体收集参数 → 延迟调度 → ProgressService 执行。
    /// </summary>
    public sealed class InsertPhotosOrchestrator
    {
        private readonly Word.Application _application;
        private readonly INotificationService _notification;

        public InsertPhotosOrchestrator(Word.Application application, INotificationService notification = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _notification = notification ?? new MessageBoxNotificationService();
        }

        public void ExecuteFromAutomationConfig()
        {
            InsertPhotosAutomationGate.EnsureEnabled();
            if (!InsertPhotosAutomationGate.TryBuildRequest(out var request))
            {
                throw new InvalidOperationException(
                    "Unable to build insert request from automation config.");
            }

            var totalFiles = request.Mode == InsertPhotosRequestMode.Folder
                ? FileService.CountTotalImageFiles(
                    request.FolderPath,
                    request.IncludeRootImages,
                    request.IncludeSubFolderImages)
                : (request.SelectedFiles?.Length ?? 0);

            var services = InsertPhotosExecutionServices.CreateHeadless(totalFiles);
            Execute(request, services);
        }

        public void ShowFormAndExecuteIfConfirmed()
        {
            try
            {
                InsertPhotosRequest pendingRequest = null;
                using (var form = new InsertPhotosForm(_application, _notification))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        pendingRequest = form.PendingRequest;
                    }
                }

                if (pendingRequest != null)
                {
                    if (InsertPhotosAutomationGate.IsEnabled)
                    {
                        Execute(pendingRequest);
                    }
                    else
                    {
                        ExecuteDeferred(pendingRequest);
                    }
                }
            }
            catch (Exception ex)
            {
                _notification.ShowError("打开窗体失败: " + ex.Message, "错误");
            }
        }

        public void Execute(InsertPhotosRequest request)
        {
            Execute(request, null);
        }

        public void Execute(InsertPhotosRequest request, InsertPhotosExecutionServices services)
        {
            if (request == null)
            {
                return;
            }

            var appContext = new WordApplicationContext(_application);

            int totalFiles = request.Mode == InsertPhotosRequestMode.Folder
                ? FileService.CountTotalImageFiles(
                    request.FolderPath,
                    request.IncludeRootImages,
                    request.IncludeSubFolderImages)
                : (request.SelectedFiles?.Length ?? 0);

            services = services ?? CreateDefaultServices(totalFiles);

            var progressService = new ProgressService(
                appContext,
                services.ProgressReporter,
                services.FailureDetailsPresenter,
                services.NotificationService);

            if (request.Mode == InsertPhotosRequestMode.Folder)
            {
                progressService.InsertPhotosWithProgress(
                    request.FolderPath,
                    request.MinHeight,
                    request.NeedDescription,
                    request.UseFileNameAsDescription,
                    request.UseFolderNameAsDescription,
                    request.IncludeRootImages,
                    request.IncludeSubFolderImages,
                    request.NeedAutoNumbering,
                    request.NumberAlignment,
                    request.NumberPosition);
                return;
            }

            progressService.InsertSelectedPhotosWithProgress(
                request.SelectedFiles,
                request.MinHeight,
                request.NeedDescription,
                request.UseFileNameAsDescription,
                request.UseFolderNameAsDescription,
                request.NeedAutoNumbering,
                request.NumberAlignment,
                request.NumberPosition);
        }

        public void ExecuteDeferred(InsertPhotosRequest request)
        {
            if (request == null)
            {
                return;
            }

            // 使用一次性 Timer 在消息循环的下一周期调度插入操作。
            // 这样可以确保模态对话框已完全关闭，避免在对话框句柄仍有效时操作 Word。
            var timer = new Timer { Interval = 1 };
            timer.Tick += (sender, e) =>
            {
                timer.Stop();
                timer.Dispose();

                try
                {
                    Execute(request);
                }
                catch (Exception ex)
                {
                    _notification.ShowError("批量插图失败: " + ex.Message, "错误");
                }
            };

            timer.Start();
        }

        private InsertPhotosExecutionServices CreateDefaultServices(int totalFiles)
        {
            return new InsertPhotosExecutionServices
            {
                ProgressReporter = new ProgressFormAdapter(totalFiles),
                NotificationService = _notification,
                FailureDetailsPresenter = new FailureDetailsFormAdapter()
            };
        }
    }
}
