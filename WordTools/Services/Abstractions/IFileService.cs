using System.Collections.Generic;

namespace WordTools.Services.Abstractions
{
    /// <summary>文件系统操作抽象。</summary>
    /// <remarks>规划抽象（Phase 2）。运行时仍使用 static FileService，尚无 Adapter 实现。</remarks>
    public interface IFileService
    {
        bool IsValidImageFile(string filePath);
        bool FileExists(string filePath);
        bool ValidateImageFile(string filePath, out string errorMessage);
        List<string> BatchValidateImageFiles(string[] filePaths, List<(string fileName, string errorReason)> failedFiles);
        string[] GetImageFiles(string folderPath, bool includeSubfolders = false);
        string[] GetRootImageFiles(string folderPath);
        string[] GetSubfolders(string folderPath);
        int CountTotalImageFiles(string folderPath, bool includeRootImages, bool includeSubFolderImages);
        string GetFileNameWithoutExtension(string filePath);
        string GetFileName(string filePath);
        string GetFolderName(string filePath);
        string GetParentFolder(string filePath);
    }
}
