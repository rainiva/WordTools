using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WordTools.Services
{
    /// <summary>
    /// 文件工具服务
    /// 处理文件选择、验证等文件操作
    /// </summary>
    public static class FileService
    {
        // 支持的图片文件扩展名
        private static readonly string[] SupportedExtensions = { ".jpg", ".jpeg", ".png" };

        #region 文件夹选择

        /// <summary>
        /// 选择文件夹
        /// </summary>
        /// <param name="dialogTitle">对话框标题</param>
        /// <param name="initialPath">初始路径（可选）</param>
        /// <returns>选中的文件夹路径，如果取消则返回空字符串</returns>
        public static string SelectFolder(string dialogTitle = "请选择文件夹...", string initialPath = "")
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = dialogTitle;
                dialog.ShowNewFolderButton = false;

                // 设置初始路径
                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    dialog.SelectedPath = initialPath;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.SelectedPath;
                }
            }
            return string.Empty;
        }

        #endregion

        #region 图片文件选择

        /// <summary>
        /// 选择图片文件（多选）
        /// </summary>
        /// <param name="dialogTitle">对话框标题</param>
        /// <param name="initialPath">初始路径（可选）</param>
        /// <returns>选中的文件路径数组，如果取消则返回 null</returns>
        public static string[] SelectImageFiles(string dialogTitle = "请选择图片文件...", string initialPath = "")
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = dialogTitle;
                dialog.Multiselect = true;
                dialog.Filter = "图片文件|*.jpg;*.jpeg;*.png|所有文件|*.*";
                dialog.FilterIndex = 1;

                // 设置初始目录
                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    dialog.InitialDirectory = initialPath;
                }

                if (dialog.ShowDialog() == DialogResult.OK && dialog.FileNames.Length > 0)
                {
                    return dialog.FileNames;
                }
            }
            return null;
        }

        #endregion

        #region 文件验证

        /// <summary>
        /// 验证是否为支持的图片文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>True 如果是支持的图片格式</returns>
        public static bool IsValidImageFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return false;

            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return SupportedExtensions.Contains(extension);
        }

        /// <summary>
        /// 验证文件是否存在
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>True 如果文件存在</returns>
        public static bool FileExists(string filePath)
        {
            return !string.IsNullOrEmpty(filePath) && File.Exists(filePath);
        }

        /// <summary>
        /// 验证图片文件是否可正常读取（存在、非零字节、有读取权限）
        /// 注意：此方法仅做轻量级检查（FileInfo），不做文件头读取以避免磁盘IO
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="errorMessage">错误信息输出</param>
        /// <returns>True 如果文件有效</returns>
        public static bool ValidateImageFile(string filePath, out string errorMessage)
        {
            errorMessage = null;

            if (string.IsNullOrEmpty(filePath))
            {
                errorMessage = "文件路径为空";
                return false;
            }

            if (!File.Exists(filePath))
            {
                errorMessage = "文件不存在";
                return false;
            }

            try
            {
                var fi = new FileInfo(filePath);
                if (fi.Length == 0)
                {
                    errorMessage = "文件大小为0字节";
                    return false;
                }
                if (fi.Length > 100 * 1024 * 1024)
                {
                    errorMessage = "文件超过100MB，无法插入";
                    return false;
                }

                return true;
            }
            catch (UnauthorizedAccessException)
            {
                errorMessage = "没有文件读取权限";
                return false;
            }
            catch (IOException ex)
            {
                errorMessage = "文件被占用: " + ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "文件验证失败: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 批量预检图片文件（在循环外调用，避免重复磁盘IO）
        /// </summary>
        /// <param name="filePaths">文件路径数组</param>
        /// <returns>有效的文件路径列表</returns>
        public static List<string> BatchValidateImageFiles(string[] filePaths, List<(string fileName, string errorReason)> failedFiles)
        {
            var validFiles = new List<string>();
            if (filePaths == null) return validFiles;

            foreach (var path in filePaths)
            {
                string errorMsg;
                if (ValidateImageFile(path, out errorMsg))
                {
                    validFiles.Add(path);
                }
                else
                {
                    failedFiles?.Add((GetFileName(path), errorMsg));
                }
            }
            return validFiles;
        }

        #endregion

        #region 文件列表获取

        /// <summary>
        /// 获取文件夹中的所有图片文件
        /// </summary>
        /// <param name="folderPath">文件夹路径</param>
        /// <param name="includeSubfolders">是否包含子文件夹</param>
        /// <returns>图片文件路径数组（已自然排序）</returns>
        public static string[] GetImageFiles(string folderPath, bool includeSubfolders = false)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                return new string[0];
            }

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = new List<string>();

            foreach (var ext in SupportedExtensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, "*" + ext, searchOption));
            }

            // 自然排序
            return NaturalSortFiles(files.ToArray());
        }

        /// <summary>
        /// 获取根目录中的图片文件（不包含子文件夹）
        /// </summary>
        public static string[] GetRootImageFiles(string folderPath)
        {
            return GetImageFiles(folderPath, false);
        }

        /// <summary>
        /// 获取子文件夹列表
        /// </summary>
        /// <param name="folderPath">父文件夹路径</param>
        /// <returns>子文件夹路径数组（已自然排序）</returns>
        public static string[] GetSubfolders(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                return new string[0];
            }

            var folders = Directory.GetDirectories(folderPath);
            return NaturalSortFolders(folders);
        }

        /// <summary>
        /// 统计总图片数量
        /// </summary>
        public static int CountTotalImageFiles(string folderPath, bool includeRootImages, bool includeSubFolderImages)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                return 0;
            }

            int count = 0;

            // 统计根目录图片
            if (includeRootImages)
            {
                count += GetRootImageFiles(folderPath).Length;
            }

            // 统计子文件夹图片
            if (includeSubFolderImages)
            {
                foreach (var subfolder in GetSubfolders(folderPath))
                {
                    count += GetRootImageFiles(subfolder).Length;
                }
            }

            return count;
        }

        /// <summary>
        /// 一次遍历获取图片文件列表和总数（避免重复扫描目录）
        /// </summary>
        public static ImageFileCollection GetImageFiles(string folderPath, bool includeRootImages, bool includeSubFolderImages)
        {
            var result = new ImageFileCollection();

            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                return result;
            }

            if (includeRootImages)
            {
                var rootFiles = GetRootImageFiles(folderPath);
                if (rootFiles != null && rootFiles.Length > 0)
                {
                    result.RootFiles = rootFiles;
                    result.TotalCount += rootFiles.Length;
                }
            }

            if (includeSubFolderImages)
            {
                var subfolders = GetSubfolders(folderPath);
                if (subfolders != null)
                {
                    foreach (var subfolder in subfolders)
                    {
                        var files = GetRootImageFiles(subfolder);
                        if (files != null && files.Length > 0)
                        {
                            result.SubfolderFiles[subfolder] = files;
                            result.TotalCount += files.Length;
                        }
                    }
                }
            }

            return result;
        }

        #endregion

        #region 图片文件集合

        /// <summary>
        /// 图片文件集合（缓存遍历结果）
        /// </summary>
        public class ImageFileCollection
        {
            public string[] RootFiles { get; set; }
            public Dictionary<string, string[]> SubfolderFiles { get; set; } = new Dictionary<string, string[]>();
            public int TotalCount { get; set; }
        }

        #endregion

        #region 自然排序

        /// <summary>
        /// 自然排序比较器（类似资源管理器）
        /// </summary>
        private static int NaturalCompare(string str1, string str2)
        {
            int i1 = 0, i2 = 0;

            while (i1 < str1.Length || i2 < str2.Length)
            {
                if (i1 >= str1.Length) return -1;
                if (i2 >= str2.Length) return 1;

                char c1 = str1[i1];
                char c2 = str2[i2];

                bool isNum1 = char.IsDigit(c1);
                bool isNum2 = char.IsDigit(c2);

                if (isNum1 && isNum2)
                {
                    // 提取完整的数字
                    string num1 = ExtractNumber(str1, ref i1);
                    string num2 = ExtractNumber(str2, ref i2);

                    // 比较数字（使用 TryParse 避免溢出风险）
                    double val1, val2;
                    bool parsed1 = double.TryParse(num1, out val1);
                    bool parsed2 = double.TryParse(num2, out val2);

                    if (parsed1 && parsed2)
                    {
                        if (val1 < val2) return -1;
                        if (val1 > val2) return 1;
                    }
                    else
                    {
                        // 解析失败时回退到字符串比较
                        int cmp = string.Compare(num1, num2, StringComparison.Ordinal);
                        if (cmp != 0) return cmp;
                    }
                }
                else
                {
                    // 按字符比较（不区分大小写）
                    int cmp = char.ToLowerInvariant(c1).CompareTo(char.ToLowerInvariant(c2));
                    if (cmp != 0) return cmp;
                    i1++;
                    i2++;
                }
            }

            return 0;
        }

        /// <summary>
        /// 从字符串中提取数字
        /// </summary>
        private static string ExtractNumber(string str, ref int index)
        {
            int start = index;
            while (index < str.Length && char.IsDigit(str[index]))
            {
                index++;
            }
            return str.Substring(start, index - start);
        }

        /// <summary>
        /// 对文件路径数组进行自然排序
        /// </summary>
        public static string[] NaturalSortFiles(string[] filePaths)
        {
            return filePaths
                .OrderBy(f => Path.GetFileName(f), Comparer<string>.Create(NaturalCompare))
                .ToArray();
        }

        /// <summary>
        /// 对文件夹路径数组进行自然排序
        /// </summary>
        public static string[] NaturalSortFolders(string[] folderPaths)
        {
            return folderPaths
                .OrderBy(f => Path.GetFileName(f), Comparer<string>.Create(NaturalCompare))
                .ToArray();
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取文件名（不含路径和扩展名）
        /// </summary>
        public static string GetFileNameWithoutExtension(string filePath)
        {
            return Path.GetFileNameWithoutExtension(filePath);
        }

        /// <summary>
        /// 获取文件名（包含扩展名）
        /// </summary>
        public static string GetFileName(string filePath)
        {
            return Path.GetFileName(filePath);
        }

        /// <summary>
        /// 获取文件夹名称
        /// </summary>
        public static string GetFolderName(string folderPath)
        {
            return new DirectoryInfo(folderPath).Name;
        }

        /// <summary>
        /// 获取父文件夹路径
        /// </summary>
        public static string GetParentFolder(string path)
        {
            return Path.GetDirectoryName(path);
        }

        #endregion
    }
}
