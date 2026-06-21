using System;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using WordTools.Services;
using WordTools.Services.Abstractions;
using WordTools.Services.Adapters;

namespace WordTools
{
    public class RibbonController
    {
        private Office.IRibbonUI _ribbonUI;
        private readonly INotificationService _notification = new MessageBoxNotificationService();

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
        }

        public void InvalidateRibbon()
        {
            _ribbonUI?.Invalidate();
        }

        public bool GetDetailedLoggingPressed(Office.IRibbonControl control)
        {
            return ConfigService.GetDetailedLoggingEnabled();
        }

        public bool GetBenchmarkLoggingPressed(Office.IRibbonControl control)
        {
            return LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(),
                ConfigService.GetBenchmarkLoggingEnabled()).BenchmarkLoggingEnabled;
        }

        public bool GetBenchmarkLoggingEnabled(Office.IRibbonControl control)
        {
            return ConfigService.GetDetailedLoggingEnabled()
                && ConfigService.GetBenchmarkLoggingEnabled();
        }

        public void OnShowLoggingSettingsSummary(Office.IRibbonControl control)
        {
            var state = LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(),
                ConfigService.GetBenchmarkLoggingEnabled());

            string message =
                "当前日志设置：\n\n" +
                "详细日志：" + (state.DetailedLoggingEnabled ? "已开启" : "已关闭") + "\n" +
                "性能基准 CSV：" + (state.BenchmarkLoggingEnabled ? "已开启" : "已关闭") + "\n\n" +
                "提示：点击右侧下拉箭头可以直接调整这两项设置。";

            _notification.ShowInformation(message, "日志设置");
        }

        public void OnInsertPhotosClick(Word.Application application)
        {
            new InsertPhotosOrchestrator(application).ShowFormAndExecuteIfConfirmed();
        }

        public void OnRefreshNumberingClick(Word.Application application)
        {
            var appContext = new WordApplicationContext(application);
            var numberingRefreshService = new NumberingRefreshService(appContext, _notification);
            numberingRefreshService.RefreshFromCurrentSelection();
        }

        public void OnToggleDetailedLogging(Office.IRibbonControl control, bool pressed)
        {
            var state = LoggingOptionsStateController.Normalize(
                pressed, ConfigService.GetBenchmarkLoggingEnabled());
            ConfigService.SaveDetailedLoggingEnabled(state.DetailedLoggingEnabled);
            ConfigService.SaveBenchmarkLoggingEnabled(state.BenchmarkLoggingEnabled);
            InvalidateRibbon();
        }

        public void OnToggleBenchmarkLogging(Office.IRibbonControl control, bool pressed)
        {
            var state = LoggingOptionsStateController.Normalize(
                ConfigService.GetDetailedLoggingEnabled(), pressed);
            ConfigService.SaveDetailedLoggingEnabled(state.DetailedLoggingEnabled);
            ConfigService.SaveBenchmarkLoggingEnabled(state.BenchmarkLoggingEnabled);
            InvalidateRibbon();
        }

        public void OnAboutClick(Office.IRibbonControl control)
        {
            _notification.ShowInformation(AppVersionInfo.AboutMessage, "关于");
        }
    }
}
