using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BatchInsertUIE2E
{
    internal enum UiFormAction
    {
        SelectFiles,
        InsertFromFolder,
    }

    internal sealed class UiCasePlan
    {
        private static readonly HashSet<string> ImageExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
        };

        public string CaseId { get; private set; }
        public UiFormAction FormAction { get; private set; }
        public int ExpectedImageCount { get; private set; }
        public TimeSpan CompletionTimeout { get; private set; }

        public static UiCasePlan Resolve(string caseId, string imageRoot)
        {
            if (string.IsNullOrWhiteSpace(caseId))
            {
                throw new ArgumentException("caseId is required", nameof(caseId));
            }

            if (string.IsNullOrWhiteSpace(imageRoot) || !Directory.Exists(imageRoot))
            {
                throw new DirectoryNotFoundException("ImageRoot folder not found: " + imageRoot);
            }

            var images = EnumerateImages(imageRoot).ToArray();
            if (images.Length == 0)
            {
                throw new InvalidOperationException("No images found under ImageRoot: " + imageRoot);
            }

            switch (caseId.ToUpperInvariant())
            {
                case "AC-UI-B03":
                    if (images.Length < 4)
                    {
                        throw new InvalidOperationException("AC-UI-B03 requires at least 4 images.");
                    }

                    return new UiCasePlan
                    {
                        CaseId = caseId,
                        FormAction = UiFormAction.SelectFiles,
                        ExpectedImageCount = 4,
                        CompletionTimeout = TimeSpan.FromMinutes(5),
                        _selectedFiles = images.Take(4).ToArray(),
                    };

                case "AC-UI-B04":
                    return new UiCasePlan
                    {
                        CaseId = caseId,
                        FormAction = UiFormAction.InsertFromFolder,
                        ExpectedImageCount = images.Length,
                        CompletionTimeout = TimeSpan.FromMinutes(20),
                        _folderPath = Path.GetFullPath(imageRoot),
                    };

                case "AC-UI-B05":
                    return new UiCasePlan
                    {
                        CaseId = caseId,
                        FormAction = UiFormAction.SelectFiles,
                        ExpectedImageCount = 1,
                        CompletionTimeout = TimeSpan.FromMinutes(3),
                        _selectedFiles = new[] { images[0] },
                    };

                default:
                    throw new NotSupportedException("Unsupported UI case: " + caseId);
            }
        }

        private string[] _selectedFiles;
        private string _folderPath;

        public void ApplyAutomationEnvironment()
        {
            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION", "1");
            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_CASE_ID", CaseId);
            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_SELECTED_FILES", null);
            Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_FOLDER_PATH", null);

            if (_selectedFiles != null && _selectedFiles.Length > 0)
            {
                Environment.SetEnvironmentVariable(
                    "WORDTOOLS_UI_AUTOMATION_SELECTED_FILES",
                    string.Join(";", _selectedFiles));
            }

            if (!string.IsNullOrWhiteSpace(_folderPath))
            {
                Environment.SetEnvironmentVariable("WORDTOOLS_UI_AUTOMATION_FOLDER_PATH", _folderPath);
            }
        }

        public void WritePlanMetadata(IDictionary<string, object> payload)
        {
            payload["expected_image_count"] = ExpectedImageCount;
            payload["form_action"] = FormAction.ToString();
            payload["image_root"] = _folderPath ?? Path.GetDirectoryName(_selectedFiles?[0]);
        }

        private static IEnumerable<string> EnumerateImages(string root)
        {
            return Directory
                .EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
                .Where(path => ImageExtensions.Contains(Path.GetExtension(path)))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);
        }
    }
}
