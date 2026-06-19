using System;

namespace WordTools.Services.Abstractions
{
    /// <summary>
    /// 进度报告抽象。由具体 UI 层实现，屏蔽服务对 ProgressForm 的直接依赖。
    /// </summary>
    public interface IProgressReporter
    {
        void Show();
        void Close();
        void UpdateProgress(int current, int total, string currentFile, TimeSpan elapsed);
        void ShowCompletion(int successCount, int failCount, double totalSeconds);
        bool IsCancelled { get; }
        IntPtr Handle { get; }
    }
}
