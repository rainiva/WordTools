using System;
using System.Runtime.InteropServices;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    public sealed class EscapeKeyMonitor
    {
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_ESCAPE = 0x1B;

        private readonly IWordApplicationContext _appContext;
        private bool _isCancelled;

        public EscapeKeyMonitor(IWordApplicationContext appContext)
        {
            _appContext = appContext ?? throw new ArgumentNullException(nameof(appContext));
        }

        public bool IsCancelled => _isCancelled;

        public bool IsEscapeKeyPressed()
        {
            return (GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0;
        }

        public bool ShouldCancel(IProgressReporter progressReporter)
        {
            if (_isCancelled) return true;

            if (progressReporter?.IsCancelled == true)
            {
                _isCancelled = true;
                return true;
            }

            if (IsEscapeKeyPressed())
            {
                _isCancelled = true;
                _appContext.Application.StatusBar = "检测到 ESC 键，正在取消操作...";
                _appContext.DoEvents();
                return true;
            }

            return false;
        }

        public void Reset()
        {
            _isCancelled = false;
        }

        public void Cancel()
        {
            _isCancelled = true;
        }
    }
}
