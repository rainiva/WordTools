using System;
using System.Linq;
using System.IO;

namespace WordTools.Services
{
    /// <summary>
    /// UI 自动化入口门禁：仅当环境变量 WORDTOOLS_UI_AUTOMATION=1 时启用测试专用 COM/表单路径。
    /// </summary>
    public static partial class InsertPhotosAutomationGate
    {
        public const string EnableEnvVar = "WORDTOOLS_UI_AUTOMATION";
        public const string SelectedFilesEnvVar = "WORDTOOLS_UI_AUTOMATION_SELECTED_FILES";
        public const string FolderPathEnvVar = "WORDTOOLS_UI_AUTOMATION_FOLDER_PATH";
        public const string ConfigFileEnvVar = "WORDTOOLS_UI_AUTOMATION_CONFIG_FILE";

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
            var raw = ReadConfigField("selected_files");
            if (string.IsNullOrWhiteSpace(raw))
            {
                raw = Environment.GetEnvironmentVariable(SelectedFilesEnvVar);
            }

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

        public static string GetPresetFolderPath()
        {
            var fromFile = ReadConfigField("folder_path");
            if (!string.IsNullOrWhiteSpace(fromFile))
            {
                return fromFile;
            }

            return Environment.GetEnvironmentVariable(FolderPathEnvVar)?.Trim();
        }

        public static bool TryGetPresetFolderPath(out string folderPath)
        {
            folderPath = GetPresetFolderPath();
            return !string.IsNullOrWhiteSpace(folderPath) && Directory.Exists(folderPath);
        }

        internal static string ReadConfigField(string fieldName)
        {
            var configPath = Environment.GetEnvironmentVariable(ConfigFileEnvVar);
            if (string.IsNullOrWhiteSpace(configPath) || !File.Exists(configPath))
            {
                return null;
            }

            try
            {
                var json = File.ReadAllText(configPath);
                var marker = "\"" + fieldName + "\"";
                var index = json.IndexOf(marker, StringComparison.Ordinal);
                if (index < 0)
                {
                    return null;
                }

                var colon = json.IndexOf(':', index);
                if (colon < 0)
                {
                    return null;
                }

                var startQuote = json.IndexOf('"', colon + 1);
                if (startQuote < 0)
                {
                    return null;
                }

                var endQuote = json.IndexOf('"', startQuote + 1);
                if (endQuote < 0)
                {
                    return null;
                }

                return json.Substring(startQuote + 1, endQuote - startQuote - 1)
                    .Replace("\\\\", "\\")
                    .Replace("\\\"", "\"");
            }
            catch
            {
                return null;
            }
        }
    }
}
