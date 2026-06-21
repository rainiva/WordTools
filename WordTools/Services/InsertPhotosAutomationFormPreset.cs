using System;

namespace WordTools.Services
{
    /// <summary>
    /// UI 自动化表单选项，与 Phase B headless E2E 及真实 Ribbon 默认流程对齐。
    /// </summary>
    public sealed class InsertPhotosAutomationFormPreset
    {
        public bool UseFileNameAsDescription { get; set; }
        public bool NeedAutoNumbering { get; set; }
        public bool IncludeRootImages { get; set; }
        public bool IncludeSubFolderImages { get; set; }
    }

    public static partial class InsertPhotosAutomationGate
    {
        public const string CaseIdEnvVar = "WORDTOOLS_UI_AUTOMATION_CASE_ID";

        public static string GetCaseId()
        {
            return Environment.GetEnvironmentVariable(CaseIdEnvVar)?.Trim();
        }

        public static InsertPhotosAutomationFormPreset ResolveFormPreset(string caseId)
        {
            var normalized = NormalizeCaseId(caseId);
            switch (normalized)
            {
                case "AC-UI-B05":
                    return new InsertPhotosAutomationFormPreset
                    {
                        UseFileNameAsDescription = true,
                        NeedAutoNumbering = false,
                        IncludeRootImages = true,
                        IncludeSubFolderImages = true,
                    };
                case "AC-UI-B04":
                    return new InsertPhotosAutomationFormPreset
                    {
                        UseFileNameAsDescription = true,
                        NeedAutoNumbering = true,
                        IncludeRootImages = true,
                        IncludeSubFolderImages = true,
                    };
                default:
                    return new InsertPhotosAutomationFormPreset
                    {
                        UseFileNameAsDescription = true,
                        NeedAutoNumbering = true,
                        IncludeRootImages = true,
                        IncludeSubFolderImages = true,
                    };
            }
        }

        public static InsertPhotosAutomationFormPreset ResolveFormPresetFromEnvironment()
        {
            return ResolveFormPreset(GetCaseId());
        }

        private static string NormalizeCaseId(string caseId)
        {
            return string.IsNullOrWhiteSpace(caseId)
                ? string.Empty
                : caseId.Trim().ToUpperInvariant();
        }
    }
}
