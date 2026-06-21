using WordTools.Forms;

namespace WordTools.Services
{
    public static partial class InsertPhotosAutomationGate
    {
        private const float CmToPoints = 28.35f;

        public static bool TryBuildRequest(out InsertPhotosRequest request)
        {
            request = null;
            var caseId = GetCaseId();
            if (string.IsNullOrWhiteSpace(caseId))
            {
                return false;
            }

            var preset = ResolveFormPreset(caseId);
            if (TryGetPresetSelectedFiles(out var selectedFiles) && selectedFiles.Length > 0)
            {
                request = CreateRequestFromPreset(preset);
                request.Mode = InsertPhotosRequestMode.SelectedFiles;
                request.SelectedFiles = selectedFiles;
                return true;
            }

            if (TryGetPresetFolderPath(out var folderPath))
            {
                request = CreateRequestFromPreset(preset);
                request.Mode = InsertPhotosRequestMode.Folder;
                request.FolderPath = folderPath;
                return true;
            }

            return false;
        }

        private static InsertPhotosRequest CreateRequestFromPreset(InsertPhotosAutomationFormPreset preset)
        {
            var minHeight = -1f;
            if (preset.MinHeightCm.HasValue && preset.MinHeightCm.Value > 0)
            {
                minHeight = preset.MinHeightCm.Value * CmToPoints;
            }

            var needDescription = preset.NeedDescription
                || preset.UseFileNameAsDescription
                || preset.UseFolderNameAsDescription;

            return new InsertPhotosRequest
            {
                MinHeight = minHeight,
                NeedDescription = needDescription,
                UseFileNameAsDescription = preset.UseFileNameAsDescription,
                UseFolderNameAsDescription = preset.UseFolderNameAsDescription,
                IncludeRootImages = preset.IncludeRootImages,
                IncludeSubFolderImages = preset.IncludeSubFolderImages,
                NeedAutoNumbering = preset.NeedAutoNumbering,
                NumberAlignment = preset.NumberAlignment,
                NumberPosition = preset.NumberPosition,
            };
        }
    }
}
