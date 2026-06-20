using System;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// 无 UI 进度报告器，供自动化 E2E 捕获完成计数与模拟取消。
    /// </summary>
    public sealed class HeadlessProgressReporter : IProgressReporter
    {
        public int LastSuccessCount { get; private set; }
        public int LastFailCount { get; private set; }
        public int CancelAfterUpdateCount { get; set; }

        private int _updateCount;
        private bool _cancelled;

        public void Show()
        {
        }

        public void Close()
        {
        }

        public void UpdateProgress(int current, int total, string currentFile, TimeSpan elapsed)
        {
            _updateCount++;
            if (CancelAfterUpdateCount > 0 && _updateCount >= CancelAfterUpdateCount)
            {
                _cancelled = true;
            }
        }

        public void ShowCompletion(int successCount, int failCount, double totalSeconds)
        {
            LastSuccessCount = successCount;
            LastFailCount = failCount;
        }

        public bool IsCancelled => _cancelled;

        public IntPtr Handle => IntPtr.Zero;
    }
}
