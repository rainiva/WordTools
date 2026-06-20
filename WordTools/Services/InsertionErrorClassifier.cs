using System;

namespace WordTools.Services
{
    public static class InsertionErrorClassifier
    {
        public static string Classify(Exception ex)
        {
            if (ex == null) return "未知错误";

            string msg = ex.Message ?? "";
            string hResult = ex.HResult.ToString("X8");

            // COM 忙碌错误
            if (msg.IndexOf("rejected", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("retry", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("busy", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("忙", StringComparison.OrdinalIgnoreCase) >= 0 ||
                ex.HResult == unchecked((int)0x80010001) ||
                ex.HResult == unchecked((int)0x8001010A))
            {
                return "Word 正忙，请关闭其他对话框后重试";
            }

            // 文件不存在
            if (ex is System.IO.FileNotFoundException ||
                msg.IndexOf("找不到", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "文件不存在或已被移动";
            }

            // 文件被占用
            if (ex is System.IO.IOException ||
                msg.IndexOf("进程无法访问", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("being used", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("占用", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "文件被其他程序占用";
            }

            // 权限错误
            if (ex is System.UnauthorizedAccessException ||
                msg.IndexOf("拒绝访问", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("access denied", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "没有文件访问权限";
            }

            // 合并单元格
            if (msg.IndexOf("合并", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("merge", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("单元格索引异常", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "目标单元格为合并单元格，无法插入";
            }

            // 行高/宽度异常
            if (msg.IndexOf("行高异常", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("宽度异常", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return msg;
            }

            // 图片格式不支持
            if (msg.IndexOf("格式", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("format", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("不支持", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("not supported", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "图片格式不受支持";
            }

            // 文件损坏
            if (msg.IndexOf("损坏", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("corrupt", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("invalid", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("0字节", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "图片文件损坏";
            }

            // 尺寸异常
            if (msg.IndexOf("尺寸异常", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("Width", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("Height", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "图片尺寸异常";
            }

            // 表格操作失败
            if (msg.IndexOf("表格", StringComparison.OrdinalIgnoreCase) >= 0 ||
                msg.IndexOf("行添加失败", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "表格操作失败: " + msg;
            }

            // 默认返回简化信息
            if (msg.Length > 80)
            {
                return msg.Substring(0, 80) + "...";
            }
            return msg;
        }

        public static bool IsMergedCellError(string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage))
            {
                return false;
            }

            return errorMessage.IndexOf("合并", StringComparison.OrdinalIgnoreCase) >= 0
                || errorMessage.IndexOf("merge", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsMergedCellError(Exception ex)
        {
            if (ex == null) return false;
            return IsMergedCellError(ex.Message);
        }
    }
}
