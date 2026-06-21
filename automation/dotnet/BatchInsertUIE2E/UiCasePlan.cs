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

        public bool ExpectZeroImages { get; private set; }

        public bool RequiresLargeBatchConfirm { get; private set; }

        public string[] SubfolderTitleHints { get; private set; } = Array.Empty<string>();



        public static UiCasePlan Resolve(string caseId, string imageRoot, string repoRoot = null)

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

            var rootImages = EnumerateImages(imageRoot, SearchOption.TopDirectoryOnly).ToArray();

            var subfolderHints = ListSubfolderNames(imageRoot);

            var subOnlyCount = images.Length - rootImages.Length;



            switch (caseId.ToUpperInvariant())

            {

                case "AC-UI-B03":

                    if (images.Length < 4)

                    {

                        throw new InvalidOperationException("AC-UI-B03 requires at least 4 images.");

                    }



                    return BuildSelected(caseId, images.Take(4).ToArray(), 4, TimeSpan.FromMinutes(4));



                case "AC-UI-B04":

                    return BuildFolder(caseId, imageRoot, images.Length, TimeSpan.FromMinutes(25), images.Length > 20, subfolderHints);



                case "AC-UI-B05":

                    return BuildSelected(caseId, new[] { images[0] }, 1, TimeSpan.FromMinutes(2));



                case "AC-UI-B07":

                    return BuildFolder(

                        caseId,

                        imageRoot,

                        rootImages.Length,

                        TimeSpan.FromMinutes(3),

                        false,

                        subfolderHints,

                        expectZeroImages: rootImages.Length == 0);



                case "AC-UI-B08":

                    if (subOnlyCount <= 0)

                    {

                        throw new InvalidOperationException("AC-UI-B08 requires subfolder images under ImageRoot.");

                    }



                    return BuildFolder(caseId, imageRoot, subOnlyCount, TimeSpan.FromMinutes(25), subOnlyCount > 20, subfolderHints);



                case "AC-UI-B09":

                case "AC-UI-B10":

                case "AC-UI-B12":

                case "AC-UI-B14":

                    if (images.Length < 4)

                    {

                        throw new InvalidOperationException(caseId + " requires at least 4 images.");

                    }



                    return BuildSelected(caseId, images.Take(4).ToArray(), 4, TimeSpan.FromMinutes(4));



                case "AC-UI-B11":

                    return BuildFolder(caseId, imageRoot, images.Length, TimeSpan.FromMinutes(25), images.Length > 20, subfolderHints);



                case "AC-UI-B15":

                    if (images.Length < 4)

                    {

                        throw new InvalidOperationException("AC-UI-B15 requires at least 4 images.");

                    }



                    return BuildSelected(caseId, images.Take(4).ToArray(), 4, TimeSpan.FromMinutes(4));



                default:

                    throw new NotSupportedException("Unsupported UI case: " + caseId);

            }

        }



        private static UiCasePlan BuildSelected(string caseId, string[] selectedFiles, int expectedCount, TimeSpan timeout)

        {

            return new UiCasePlan

            {

                CaseId = caseId,

                FormAction = UiFormAction.SelectFiles,

                ExpectedImageCount = expectedCount,

                CompletionTimeout = timeout,

                _selectedFiles = selectedFiles,

            };

        }



        private static UiCasePlan BuildFolder(

            string caseId,

            string folderPath,

            int expectedCount,

            TimeSpan timeout,

            bool requiresConfirm,

            string[] subfolderHints,

            bool expectZeroImages = false)

        {

            return new UiCasePlan

            {

                CaseId = caseId,

                FormAction = UiFormAction.InsertFromFolder,

                ExpectedImageCount = expectedCount,

                CompletionTimeout = timeout,

                RequiresLargeBatchConfirm = requiresConfirm,

                SubfolderTitleHints = subfolderHints ?? Array.Empty<string>(),

                ExpectZeroImages = expectZeroImages,

                _folderPath = Path.GetFullPath(folderPath),

            };

        }



        private static string[] ListSubfolderNames(string imageRoot)

        {

            return Directory

                .GetDirectories(imageRoot)

                .Select(Path.GetFileName)

                .Where(name => !string.IsNullOrWhiteSpace(name))

                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)

                .ToArray();

        }



        private static IEnumerable<string> EnumerateImages(string root, SearchOption searchOption = SearchOption.AllDirectories)

        {

            if (!Directory.Exists(root))

            {

                return Enumerable.Empty<string>();

            }



            return Directory

                .EnumerateFiles(root, "*.*", searchOption)

                .Where(path => ImageExtensions.Contains(Path.GetExtension(path)))

                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);

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



        public void WriteAutomationConfigFile(string configPath)
        {
            if (string.IsNullOrWhiteSpace(configPath))
            {
                return;
            }

            var parts = new List<string>
            {
                "\"case_id\":\"" + EscapeJson(CaseId) + "\"",
            };

            if (_selectedFiles != null && _selectedFiles.Length > 0)
            {
                parts.Add("\"selected_files\":\"" + EscapeJson(string.Join(";", _selectedFiles)) + "\"");
            }

            if (!string.IsNullOrWhiteSpace(_folderPath))
            {
                parts.Add("\"folder_path\":\"" + EscapeJson(_folderPath) + "\"");
            }

            File.WriteAllText(configPath, "{" + string.Join(",", parts) + "}");
        }

        private static string EscapeJson(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"");
        }



        public void WritePlanMetadata(IDictionary<string, object> payload)

        {

            payload["expected_image_count"] = ExpectedImageCount;

            payload["form_action"] = FormAction.ToString();

            payload["image_root"] = _folderPath ?? Path.GetDirectoryName(_selectedFiles?[0]);

            payload["expect_zero_images"] = ExpectZeroImages;

            payload["requires_large_batch_confirm"] = RequiresLargeBatchConfirm;

            payload["subfolder_title_hints"] = SubfolderTitleHints;

        }

    }

}


