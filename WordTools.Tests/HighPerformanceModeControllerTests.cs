using WordTools.Services;
using WordTools.Services.Abstractions;
using Xunit;

namespace WordTools.Tests
{
    public class HighPerformanceModeControllerTests
    {
        [Theory]
        [InlineData(10, 10)]
        [InlineData(29, 10)]
        [InlineData(30, 15)]
        [InlineData(99, 15)]
        [InlineData(100, 20)]
        [InlineData(500, 20)]
        public void GetOptimizedRefreshInterval_ReturnsExpected(int totalFiles, int expected)
        {
            var controller = new HighPerformanceModeController(new FakeAppContext());
            Assert.Equal(expected, controller.GetOptimizedRefreshInterval(totalFiles));
        }

        [Theory]
        [InlineData(5, 1)]
        [InlineData(10, 1)]
        [InlineData(11, 5)]
        [InlineData(50, 5)]
        [InlineData(51, 15)]
        [InlineData(200, 15)]
        [InlineData(201, 25)]
        public void GetStatusBarUpdateInterval_ReturnsExpected(int totalFiles, int expected)
        {
            var controller = new HighPerformanceModeController(new FakeAppContext());
            Assert.Equal(expected, controller.GetStatusBarUpdateInterval(totalFiles));
        }
    }

    internal class FakeAppContext : IWordApplicationContext
    {
        public Microsoft.Office.Interop.Word.Application Application => null;
        public bool ScreenUpdating { get; set; } = true;
        public void SetStatusBar(string text) { }
        public void DoEvents() { }
    }
}
