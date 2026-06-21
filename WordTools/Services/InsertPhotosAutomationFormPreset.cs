using System;



namespace WordTools.Services

{

    /// <summary>

    /// UI 自动化表单选项，与 Phase B headless E2E 及真实 Ribbon 默认流程对齐。

    /// </summary>

    public sealed class InsertPhotosAutomationFormPreset

    {

        public bool UseFileNameAsDescription { get; set; }

        public bool UseFolderNameAsDescription { get; set; }

        public bool NeedDescription { get; set; }

        public bool NeedAutoNumbering { get; set; }

        public bool IncludeRootImages { get; set; } = true;

        public bool IncludeSubFolderImages { get; set; } = true;

        public int NumberAlignment { get; set; } = 2;

        public int NumberPosition { get; set; } = 1;

        /// <summary>最小高度（cm）；null 表示留空、按单元格自适应。</summary>

        public float? MinHeightCm { get; set; }

    }



    public static partial class InsertPhotosAutomationGate

    {

        public const string CaseIdEnvVar = "WORDTOOLS_UI_AUTOMATION_CASE_ID";



        public static string GetCaseId()
        {
            var fromFile = ReadConfigField("case_id");
            if (!string.IsNullOrWhiteSpace(fromFile))
            {
                return fromFile.Trim();
            }

            return Environment.GetEnvironmentVariable(CaseIdEnvVar)?.Trim();
        }



        public static InsertPhotosAutomationFormPreset ResolveFormPreset(string caseId)

        {

            var normalized = NormalizeCaseId(caseId);

            switch (normalized)

            {

                case "AC-UI-B05":

                    return FilenamePreset(needAutoNumbering: false);

                case "AC-UI-B07":

                    return FilenamePreset(includeRootImages: true, includeSubFolderImages: false);

                case "AC-UI-B08":

                    return FilenamePreset(includeRootImages: false, includeSubFolderImages: true);

                case "AC-UI-B09":

                    return NoDescriptionPreset();

                case "AC-UI-B10":

                    return FilenamePreset(numberAlignment: 2, numberPosition: 2);

                case "AC-UI-B11":

                    return FolderNamePreset();

                case "AC-UI-B12":

                    return ManualDescriptionPreset();

                case "AC-UI-B14":

                    return FilenamePreset(numberAlignment: 1, numberPosition: 1);

                case "AC-UI-B15":

                    return FilenamePreset(minHeightCm: 3f);

                case "AC-UI-B04":

                    return FilenamePreset();

                default:

                    return FilenamePreset();

            }

        }



        public static InsertPhotosAutomationFormPreset ResolveFormPresetFromEnvironment()

        {

            return ResolveFormPreset(GetCaseId());

        }



        private static InsertPhotosAutomationFormPreset FilenamePreset(

            bool needAutoNumbering = true,

            bool includeRootImages = true,

            bool includeSubFolderImages = true,

            int numberAlignment = 2,

            int numberPosition = 1,

            float? minHeightCm = null)

        {

            return new InsertPhotosAutomationFormPreset

            {

                UseFileNameAsDescription = true,

                NeedAutoNumbering = needAutoNumbering,

                IncludeRootImages = includeRootImages,

                IncludeSubFolderImages = includeSubFolderImages,

                NumberAlignment = numberAlignment,

                NumberPosition = numberPosition,

                MinHeightCm = minHeightCm,

            };

        }



        private static InsertPhotosAutomationFormPreset NoDescriptionPreset()

        {

            return new InsertPhotosAutomationFormPreset

            {

                UseFileNameAsDescription = false,

                UseFolderNameAsDescription = false,

                NeedDescription = false,

                NeedAutoNumbering = false,

                IncludeRootImages = true,

                IncludeSubFolderImages = true,

            };

        }



        private static InsertPhotosAutomationFormPreset FolderNamePreset()

        {

            return new InsertPhotosAutomationFormPreset

            {

                UseFileNameAsDescription = false,

                UseFolderNameAsDescription = true,

                NeedAutoNumbering = true,

                IncludeRootImages = true,

                IncludeSubFolderImages = true,

                NumberAlignment = 2,

                NumberPosition = 1,

            };

        }



        private static InsertPhotosAutomationFormPreset ManualDescriptionPreset()

        {

            return new InsertPhotosAutomationFormPreset

            {

                NeedDescription = true,

                UseFileNameAsDescription = false,

                UseFolderNameAsDescription = false,

                NeedAutoNumbering = true,

                NumberAlignment = 2,

                NumberPosition = 1,

            };

        }



        private static string NormalizeCaseId(string caseId)

        {

            return string.IsNullOrWhiteSpace(caseId)

                ? string.Empty

                : caseId.Trim().ToUpperInvariant();

        }

    }

}


