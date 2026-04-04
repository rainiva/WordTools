using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace WordTools.Services
{
    /// <summary>
    /// 配置管理服务
    /// 管理文档自定义属性和应用程序设置的保存和读取
    /// </summary>
    public static class ConfigService
    {
        // 配置键名常量
        private const string CONFIG_KEY_LAST_IMAGE_HEIGHT_CM = "LastImageHeightCM";
        private const string CONFIG_KEY_LAST_FOLDER_PATH = "LastFolderPath";
        private const string CONFIG_KEY_NEED_DESCRIPTION = "NeedDescription";
        private const string CONFIG_KEY_USE_FILENAME_AS_DESCRIPTION = "UseFilenameAsDescription";
        private const string CONFIG_KEY_INCLUDE_ROOT_IMAGES = "IncludeRootImages";
        private const string CONFIG_KEY_INCLUDE_SUBFOLDER_IMAGES = "IncludeSubFolderImages";
        private const string CONFIG_KEY_AUTO_NUMBERING = "AutoNumbering";
        private const string CONFIG_KEY_NUMBER_ALIGNMENT = "NumberAlignment";

        // 注册表路径
        private const string REGISTRY_PATH = @"Software\WordTools";

        #region 文档属性操作

        /// <summary>
        /// 从文档自定义属性读取值
        /// </summary>
        private static string GetDocumentProperty(Document doc, string propertyName, string defaultValue = "")
        {
            if (doc == null) return defaultValue;

            try
            {
                dynamic properties = doc.CustomDocumentProperties;
                int count = properties.Count;
                
                for (int i = 1; i <= count; i++)
                {
                    dynamic prop = properties.Item(i);
                    if ((string)prop.Name == propertyName)
                    {
                        return prop.Value != null ? prop.Value.ToString() : defaultValue;
                    }
                }
            }
            catch
            {
                // 属性不存在或读取错误
            }
            return defaultValue;
        }

        /// <summary>
        /// 保存值到文档自定义属性
        /// </summary>
        private static void SetDocumentProperty(Document doc, string propertyName, string value)
        {
            if (doc == null) return;

            try
            {
                dynamic properties = doc.CustomDocumentProperties;
                bool found = false;
                int count = properties.Count;

                // 尝试更新现有属性
                for (int i = 1; i <= count; i++)
                {
                    dynamic prop = properties.Item(i);
                    if ((string)prop.Name == propertyName)
                    {
                        prop.Value = value;
                        found = true;
                        break;
                    }
                }

                // 如果属性不存在，添加新属性
                // msoPropertyTypeString = 4
                if (!found)
                {
                    properties.Add(propertyName, false, 4, value);
                }
            }
            catch
            {
                // 忽略保存错误
            }
        }

        #endregion

        #region 注册表操作

        /// <summary>
        /// 从注册表读取值
        /// </summary>
        private static string GetRegistryValue(string keyName, string defaultValue = "")
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        var value = key.GetValue(keyName, defaultValue);
                        return value != null ? value.ToString() : defaultValue;
                    }
                    return defaultValue;
                }
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 保存值到注册表
        /// </summary>
        private static void SetRegistryValue(string keyName, string value)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        key.SetValue(keyName, value);
                    }
                }
            }
            catch
            {
                // 忽略保存错误
            }
        }

        #endregion

        #region 图片高度配置

        /// <summary>
        /// 获取最后保存的图片高度（厘米）
        /// </summary>
        public static string GetLastImageHeightCM(Document doc = null)
        {
            string value = null;
            // 优先从文档属性读取
            if (doc != null)
            {
                value = GetDocumentProperty(doc, CONFIG_KEY_LAST_IMAGE_HEIGHT_CM);
                if (!string.IsNullOrEmpty(value))
                {
                    return value == "__EMPTY__" ? "" : value;
                }
            }
            // 回退到注册表
            value = GetRegistryValue(CONFIG_KEY_LAST_IMAGE_HEIGHT_CM, "");
            return value == "__EMPTY__" ? "" : value;
        }

        /// <summary>
        /// 保存图片高度（厘米）
        /// </summary>
        public static void SaveLastImageHeightCM(string heightCM, Document doc = null)
        {
            // 空值使用特殊标记保存，避免 Word COM 不接受空字符串
            string saveValue = string.IsNullOrEmpty(heightCM) ? "__EMPTY__" : heightCM;
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_LAST_IMAGE_HEIGHT_CM, saveValue);
            }
            SetRegistryValue(CONFIG_KEY_LAST_IMAGE_HEIGHT_CM, saveValue);
        }

        #endregion

        #region 文件夹路径配置

        /// <summary>
        /// 获取最后保存的文件夹路径
        /// </summary>
        public static string GetLastFolderPath(Document doc = null)
        {
            if (doc != null)
            {
                var value = GetDocumentProperty(doc, CONFIG_KEY_LAST_FOLDER_PATH);
                if (!string.IsNullOrEmpty(value)) return value;
            }
            return GetRegistryValue(CONFIG_KEY_LAST_FOLDER_PATH, "");
        }

        /// <summary>
        /// 保存文件夹路径
        /// </summary>
        public static void SaveLastFolderPath(string folderPath, Document doc = null)
        {
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_LAST_FOLDER_PATH, folderPath);
            }
            SetRegistryValue(CONFIG_KEY_LAST_FOLDER_PATH, folderPath);
        }

        #endregion

        #region 描述行配置

        /// <summary>
        /// 获取是否需要描述行
        /// </summary>
        public static bool GetNeedDescription(Document doc = null)
        {
            var value = doc != null 
                ? GetDocumentProperty(doc, CONFIG_KEY_NEED_DESCRIPTION) 
                : GetRegistryValue(CONFIG_KEY_NEED_DESCRIPTION);
            return value == "True";
        }

        /// <summary>
        /// 保存是否需要描述行
        /// </summary>
        public static void SaveNeedDescription(bool value, Document doc = null)
        {
            var strValue = value ? "True" : "False";
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_NEED_DESCRIPTION, strValue);
            }
            SetRegistryValue(CONFIG_KEY_NEED_DESCRIPTION, strValue);
        }

        /// <summary>
        /// 获取是否使用文件名作为描述
        /// </summary>
        public static bool GetUseFilenameAsDescription(Document doc = null)
        {
            var value = doc != null 
                ? GetDocumentProperty(doc, CONFIG_KEY_USE_FILENAME_AS_DESCRIPTION) 
                : GetRegistryValue(CONFIG_KEY_USE_FILENAME_AS_DESCRIPTION);
            return value == "True";
        }

        /// <summary>
        /// 保存是否使用文件名作为描述
        /// </summary>
        public static void SaveUseFilenameAsDescription(bool value, Document doc = null)
        {
            var strValue = value ? "True" : "False";
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_USE_FILENAME_AS_DESCRIPTION, strValue);
            }
            SetRegistryValue(CONFIG_KEY_USE_FILENAME_AS_DESCRIPTION, strValue);
        }

        #endregion

        #region 文件范围配置

        /// <summary>
        /// 获取是否包含根目录图片
        /// </summary>
        public static bool GetIncludeRootImages(Document doc = null)
        {
            var value = doc != null 
                ? GetDocumentProperty(doc, CONFIG_KEY_INCLUDE_ROOT_IMAGES) 
                : GetRegistryValue(CONFIG_KEY_INCLUDE_ROOT_IMAGES);
            return string.IsNullOrEmpty(value) || value == "True"; // 默认为 true
        }

        /// <summary>
        /// 保存是否包含根目录图片
        /// </summary>
        public static void SaveIncludeRootImages(bool value, Document doc = null)
        {
            var strValue = value ? "True" : "False";
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_INCLUDE_ROOT_IMAGES, strValue);
            }
            SetRegistryValue(CONFIG_KEY_INCLUDE_ROOT_IMAGES, strValue);
        }

        /// <summary>
        /// 获取是否包含子目录图片
        /// </summary>
        public static bool GetIncludeSubFolderImages(Document doc = null)
        {
            var value = doc != null 
                ? GetDocumentProperty(doc, CONFIG_KEY_INCLUDE_SUBFOLDER_IMAGES) 
                : GetRegistryValue(CONFIG_KEY_INCLUDE_SUBFOLDER_IMAGES);
            return string.IsNullOrEmpty(value) || value == "True"; // 默认为 true
        }

        /// <summary>
        /// 保存是否包含子目录图片
        /// </summary>
        public static void SaveIncludeSubFolderImages(bool value, Document doc = null)
        {
            var strValue = value ? "True" : "False";
            if (doc != null)
            {
                SetDocumentProperty(doc, CONFIG_KEY_INCLUDE_SUBFOLDER_IMAGES, strValue);
            }
            SetRegistryValue(CONFIG_KEY_INCLUDE_SUBFOLDER_IMAGES, strValue);
        }

        #endregion

        #region 自动编号配置

        /// <summary>
        /// 获取是否启用自动编号
        /// </summary>
        public static bool GetAutoNumbering()
        {
            return GetRegistryValue(CONFIG_KEY_AUTO_NUMBERING) == "True";
        }

        /// <summary>
        /// 保存是否启用自动编号
        /// </summary>
        public static void SaveAutoNumbering(bool value)
        {
            SetRegistryValue(CONFIG_KEY_AUTO_NUMBERING, value ? "True" : "False");
        }

        #endregion

        #region 编号对齐配置

        /// <summary>
        /// 获取编号对齐方式 (1=靠左, 2=居中，默认2)
        /// </summary>
        public static int GetNumberAlignment()
        {
            var value = GetRegistryValue(CONFIG_KEY_NUMBER_ALIGNMENT, "2");
            int result;
            if (int.TryParse(value, out result))
            {
                return result == 1 ? 1 : 2; // 只有1或2有效，其他值默认为2
            }
            return 2; // 默认居中
        }

        /// <summary>
        /// 保存编号对齐方式
        /// </summary>
        public static void SaveNumberAlignment(int alignment)
        {
            SetRegistryValue(CONFIG_KEY_NUMBER_ALIGNMENT, alignment.ToString());
        }

        #endregion
    }
}
