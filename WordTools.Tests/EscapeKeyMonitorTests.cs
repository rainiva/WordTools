using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class EscapeKeyMonitorTests
    {
        [Fact]
        public void IsCancelled_InitiallyFalse()
        {
            var monitor = new EscapeKeyMonitor(new FakeAppContext());
            Assert.False(monitor.IsCancelled);
        }

        [Fact]
        public void Reset_ClearsCancelledState()
        {
            var monitor = new EscapeKeyMonitor(new FakeAppContext());
            monitor.Cancel();
            monitor.Reset();
            Assert.False(monitor.IsCancelled);
        }
    }
}
