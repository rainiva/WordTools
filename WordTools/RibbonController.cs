using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using WordTools.Services;

namespace WordTools
{
    public class RibbonController
    {
        private Office.IRibbonUI _ribbonUI;

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

            MessageBox.Show(message, "日志设置",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            MessageBox.Show(AppVersionInfo.AboutMessage, "关于",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
