using Microsoft.Office.Interop.Word;

namespace WordTools.Services.Abstractions
{
    public interface IConfigService
    {
        string GetLastImageHeightCM(Document doc = null);
        void SaveLastImageHeightCM(string heightCM, Document doc = null);
        string GetLastFolderPath(Document doc = null);
        void SaveLastFolderPath(string folderPath, Document doc = null);
        bool GetNeedDescription(Document doc = null);
        void SaveNeedDescription(bool value, Document doc = null);
        bool GetUseFilenameAsDescription(Document doc = null);
        void SaveUseFilenameAsDescription(bool value, Document doc = null);
        bool GetUseFolderNameAsDescription(Document doc = null);
        void SaveUseFolderNameAsDescription(bool value, Document doc = null);
        bool GetIncludeRootImages(Document doc = null);
        void SaveIncludeRootImages(bool value, Document doc = null);
        bool GetIncludeSubFolderImages(Document doc = null);
        void SaveIncludeSubFolderImages(bool value, Document doc = null);
        bool GetAutoNumbering(Document doc = null);
        void SaveAutoNumbering(bool value, Document doc = null);
        int GetNumberAlignment(Document doc = null);
        void SaveNumberAlignment(int value, Document doc = null);
        int GetNumberPosition(Document doc = null);
        void SaveNumberPosition(int value, Document doc = null);
        bool GetDetailedLoggingEnabled();
        void SaveDetailedLoggingEnabled(bool value);
        bool GetBenchmarkLoggingEnabled();
        void SaveBenchmarkLoggingEnabled(bool value);
    }
}
