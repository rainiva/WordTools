using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WordTools.Services
{
    public static class WindowActivationService
    {
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        public static void EnsureWindowTopMost(IntPtr hWnd)
        {
            try
            {
                if (hWnd != IntPtr.Zero && IsWindow(hWnd))
                {
                    SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }
            catch (Exception ex) { Debug.WriteLine($"设置窗口置顶失败: {ex.Message}"); }
        }

        public static void EnsureWordWindowActive(IntPtr wordHandle)
        {
            try
            {
                if (wordHandle != IntPtr.Zero && IsWindow(wordHandle))
                {
                    SetWindowPos(wordHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }
            catch (Exception ex) { Debug.WriteLine($"激活 Word 窗口失败: {ex.Message}"); }
        }
    }
}
