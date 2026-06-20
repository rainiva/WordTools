using System;
using WordTools.Services;
using Xunit;

namespace WordTools.Tests
{
    public class WindowActivationServiceTests
    {
        [Fact]
        public void EnsureWindowTopMost_ZeroHandle_DoesNotThrow()
        {
            WindowActivationService.EnsureWindowTopMost(IntPtr.Zero);
        }

        [Fact]
        public void EnsureWordWindowActive_ZeroHandle_DoesNotThrow()
        {
            WindowActivationService.EnsureWordWindowActive(IntPtr.Zero);
        }
    }
}
