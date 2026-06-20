using WordTools.Services.Abstractions;
using WordTools.Services.Adapters;

namespace WordTools.Services
{
    /// <summary>
    /// 批量插图 Execute 的可注入依赖集合；生产默认仍使用 WinForms + MessageBox。
    /// </summary>
    public sealed class InsertPhotosExecutionServices
    {
        public IProgressReporter ProgressReporter { get; set; }
        public INotificationService NotificationService { get; set; }
        public IFailureDetailsPresenter FailureDetailsPresenter { get; set; }

        public static InsertPhotosExecutionServices CreateHeadless(int totalFiles, int cancelAfterUpdateCount = 0)
        {
            return new InsertPhotosExecutionServices
            {
                ProgressReporter = new HeadlessProgressReporter { CancelAfterUpdateCount = cancelAfterUpdateCount },
                NotificationService = new CapturingNotificationService(),
                FailureDetailsPresenter = new NullFailureDetailsPresenter()
            };
        }
    }
}
