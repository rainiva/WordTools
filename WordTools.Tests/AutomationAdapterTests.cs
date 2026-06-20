using System.Linq;
using System.Windows.Forms;
using WordTools.Services;
using WordTools.Services.Adapters;
using Xunit;

namespace WordTools.Tests
{
    public class AutomationAdapterTests
    {
        [Fact]
        public void HeadlessProgressReporter_ShowCompletion_RecordsCounts()
        {
            var reporter = new HeadlessProgressReporter();

            reporter.ShowCompletion(4, 0, 1.5);

            Assert.Equal(4, reporter.LastSuccessCount);
            Assert.Equal(0, reporter.LastFailCount);
        }

        [Fact]
        public void HeadlessProgressReporter_CancelAfter_ReportsCancelled()
        {
            var reporter = new HeadlessProgressReporter { CancelAfterUpdateCount = 1 };

            reporter.Show();
            reporter.UpdateProgress(1, 4, "01.jpg", System.TimeSpan.FromSeconds(1));

            Assert.True(reporter.IsCancelled);
        }

        [Fact]
        public void CapturingNotificationService_ShowWarning_RecordsMessage()
        {
            var notifications = new CapturingNotificationService();

            notifications.ShowWarning("请先选中一个表格！", "提示");

            Assert.Contains("请先选中一个表格！", notifications.Warnings.Select(w => w.Message));
        }

        [Fact]
        public void CapturingNotificationService_ShowQuestion_ReturnsYesByDefault()
        {
            var notifications = new CapturingNotificationService();

            var result = notifications.ShowQuestion("继续?", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            Assert.Equal(DialogResult.Yes, result);
        }

        [Fact]
        public void NullFailureDetailsPresenter_ShowDetails_ReturnsFalse()
        {
            var presenter = new NullFailureDetailsPresenter();

            bool copyRequested = presenter.ShowDetails("summary", null, null, null);

            Assert.False(copyRequested);
        }

        [Fact]
        public void InsertPhotosExecutionServices_CreateHeadless_UsesCapturingNotificationService()
        {
            var services = InsertPhotosExecutionServices.CreateHeadless(4);

            Assert.IsType<HeadlessProgressReporter>(services.ProgressReporter);
            Assert.IsType<CapturingNotificationService>(services.NotificationService);
            Assert.IsType<NullFailureDetailsPresenter>(services.FailureDetailsPresenter);
        }
    }
}
