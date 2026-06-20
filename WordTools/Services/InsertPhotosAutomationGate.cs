using System;
using System.Linq;

namespace WordTools.Services
{
    /// <summary>
    /// UI 自动化入口门禁：仅当环境变量 WORDTOOLS_UI_AUTOMATION=1 时启用测试专用 COM/表单路径。
    /// </summary>
    public static class InsertPhotosAutomationGate
    {
        public const string EnableEnvVar = "WORDTOOLS_UI_AUTOMATION";
        public const string SelectedFilesEnvVar = "WORDTOOLS_UI_AUTOMATION_SELECTED_FILES";

        public static bool IsEnabled =>
            string.Equals(
                Environment.GetEnvironmentVariable(EnableEnvVar),
                "1",
                StringComparison.Ordinal);

        public static void EnsureEnabled()
        {
            if (!IsEnabled)
            {
                throw new InvalidOperationException(
                    "UI automation entry is disabled. Set WORDTOOLS_UI_AUTOMATION=1 for test runs only.");
            }
        }

        public static string[] GetPresetSelectedFiles()
        {
            var raw = Environment.GetEnvironmentVariable(SelectedFilesEnvVar);
            if (string.IsNullOrWhiteSpace(raw))
            {
                return Array.Empty<string>();
            }

            return raw
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(path => path.Trim())
                .Where(path => path.Length > 0)
                .ToArray();
        }

        public static bool TryGetPresetSelectedFiles(out string[] files)
        {
            files = GetPresetSelectedFiles();
            return files.Length > 0;
        }
    }
}
