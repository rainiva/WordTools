using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using WordTools.Services.Abstractions;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordTools.Services
{
    public sealed class HighPerformanceModeController
    {
        private readonly IWordApplicationContext _appContext;
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;
        private bool _highPerformanceModeEntered;

        public HighPerformanceModeController(IWordApplicationContext appContext)
        {
            _appContext = appContext ?? throw new ArgumentNullException(nameof(appContext));
        }

        public void Enter()
        {
            if (_highPerformanceModeEntered)
            {
                return;
            }

            try
            {
                _originalScreenUpdating = _appContext.Application.ScreenUpdating;
                _originalDisplayAlerts = _appContext.Application.DisplayAlerts != WdAlertLevel.wdAlertsNone;

                // 关闭 ScreenUpdating 以提升插图性能（进度由独立窗口显示）
                _appContext.Application.ScreenUpdating = false;
                _appContext.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                var doc = _appContext.Application.ActiveDocument;
                if (doc != null)
                {
                    doc.SpellingChecked = true;
                    doc.GrammarChecked = true;
                }

                _highPerformanceModeEntered = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[HighPerformanceModeController] Enter error: {ex.Message}");
            }
        }

        /// <summary>
        /// 退出高性能模式。仅在成功进入过时执行恢复，避免取消/验证失败路径把 Word 的
        /// ScreenUpdating/DisplayAlerts 恢复到错误默认值，导致光标或界面异常。
        /// </summary>
        public void Exit()
        {
            if (!_highPerformanceModeEntered)
            {
                return;
            }

            try
            {
                _appContext.Application.ScreenUpdating = _originalScreenUpdating;
                _appContext.Application.DisplayAlerts = _originalDisplayAlerts 
                    ? WdAlertLevel.wdAlertsAll 
                    : WdAlertLevel.wdAlertsNone;
                _highPerformanceModeEntered = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"退出高性能模式失败: {ex.Message}");
            }
        }

        public int GetOptimizedRefreshInterval(int totalFiles)
        {
            if (totalFiles < 30) return 10;
            if (totalFiles < 100) return 15;
            return 20;
        }

        public int GetStatusBarUpdateInterval(int totalFiles)
        {
            if (totalFiles <= 10) return 1;     // 10 张以内：每张都更新
            if (totalFiles <= 50) return 5;     // 50 张以内：每 5 张更新
            if (totalFiles <= 200) return 15;   // 200 张以内：每 15 张更新
            return 25;                           // 更多：每 25 张更新（最大化性能）
        }
    }
}
