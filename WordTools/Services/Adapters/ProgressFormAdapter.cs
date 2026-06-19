using System;
using WordTools.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services.Adapters
{
    /// <summary>
    /// IProgressReporter 的 WinForms 实现，封装 ProgressForm。
    /// </summary>
    public sealed class ProgressFormAdapter : IProgressReporter, IDisposable
    {
        private ProgressForm _form;
        private readonly int _totalFiles;

        public ProgressFormAdapter(int totalFiles)
        {
            _totalFiles = totalFiles;
        }

        public void Show()
        {
            if (_form == null || _form.IsDisposed)
            {
                _form = new ProgressForm(_totalFiles);
            }
            _form.Show();
        }

        public void Close()
        {
            if (_form != null && !_form.IsDisposed)
            {
                _form.TopMost = false;
                _form.Close();
                _form.Dispose();
                _form = null;
            }
        }

        public void UpdateProgress(int current, int total, string currentFile, TimeSpan elapsed)
        {
            _form?.UpdateProgress(current, total, currentFile, elapsed);
        }

        public void ShowCompletion(int successCount, int failCount, double totalSeconds)
        {
            _form?.ShowCompletion(successCount, failCount, totalSeconds);
        }

        public bool IsCancelled => _form != null && !_form.IsDisposed && _form.IsCancelled;

        public IntPtr Handle => _form != null && !_form.IsDisposed ? _form.Handle : IntPtr.Zero;

        public void Dispose()
        {
            Close();
        }
    }
}
