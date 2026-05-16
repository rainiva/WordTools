using System;
using System.Collections.Generic;
using System.Text;

namespace WordTools.Services
{
    public static class InsertionSummaryFormatter
    {
        public static bool ShouldShowSummary(
            int failCount,
            IList<int> mergedCellRows,
            IList<string> overwriteWarnings)
        {
            return failCount > 0
                || (mergedCellRows != null && mergedCellRows.Count > 0)
                || (overwriteWarnings != null && overwriteWarnings.Count > 0);
        }

        public static bool HasMoreDetails(
            int previewCount,
            IList<(string fileName, string errorReason)> failedFiles,
            IList<int> mergedCellRows,
            IList<string> overwriteWarnings)
        {
            return GetCount(failedFiles) > previewCount
                || GetCount(mergedCellRows) > previewCount
                || GetCount(overwriteWarnings) > previewCount;
        }

        public static string BuildSummaryMessage(
            int successCount,
            int failCount,
            string timeInfo,
            string timeDetail,
            IList<(string fileName, string errorReason)> failedFiles,
            IList<int> mergedCellRows,
            IList<string> overwriteWarnings,
            int previewCount = 5)
        {
            var sb = new StringBuilder();

            failedFiles = failedFiles ?? new List<(string fileName, string errorReason)>();
            mergedCellRows = mergedCellRows ?? new List<int>();
            overwriteWarnings = overwriteWarnings ?? new List<string>();

            sb.AppendLine("图片插入完成：");
            sb.AppendLine(string.Format("成功: {0} 张", successCount));
            sb.AppendLine(string.Format("失败: {0} 张", failCount));

            if (mergedCellRows.Count > 0)
            {
                sb.AppendLine(string.Format("合并单元格绕开: {0} 处", mergedCellRows.Count));
            }

            if (overwriteWarnings.Count > 0)
            {
                sb.AppendLine(string.Format("覆盖插图提示: {0} 处", overwriteWarnings.Count));
            }

            sb.AppendLine(string.Format("耗时: {0}", timeInfo ?? string.Empty));

            if (failCount > 0 || failedFiles.Count > 0)
            {
                AppendFailureSection(sb, failedFiles, previewCount);
            }
            AppendMergedSection(sb, mergedCellRows, previewCount);
            AppendOverwriteSection(sb, overwriteWarnings, previewCount);

            if (!string.IsNullOrEmpty(timeDetail))
            {
                sb.Append(timeDetail);
            }

            return sb.ToString();
        }

        public static string BuildDetailsPrompt()
        {
            return "选择“是(Y)”查看详情，选择“否(N)”关闭。";
        }

        private static void AppendFailureSection(
            StringBuilder sb,
            IList<(string fileName, string errorReason)> failedFiles,
            int previewCount)
        {
            sb.AppendLine();
            sb.AppendLine("失败详情（前5项）：");

            int showCount = Math.Min(failedFiles.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine(string.Format("  {0}: {1}", failedFiles[i].fileName, failedFiles[i].errorReason));
            }

            if (failedFiles.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 个文件失败", failedFiles.Count - previewCount));
            }
        }

        private static void AppendMergedSection(StringBuilder sb, IList<int> mergedCellRows, int previewCount)
        {
            if (mergedCellRows.Count == 0)
            {
                return;
            }

            sb.AppendLine();
            sb.AppendLine("合并单元格位置（已自动绕开）：");

            int showCount = Math.Min(mergedCellRows.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine(string.Format("  第{0}行", mergedCellRows[i]));
            }

            if (mergedCellRows.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 处", mergedCellRows.Count - previewCount));
            }
        }

        private static void AppendOverwriteSection(StringBuilder sb, IList<string> overwriteWarnings, int previewCount)
        {
            if (overwriteWarnings.Count == 0)
            {
                return;
            }

            sb.AppendLine();
            sb.AppendLine("覆盖插图提示（前5项）：");

            int showCount = Math.Min(overwriteWarnings.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine("  " + overwriteWarnings[i]);
            }

            if (overwriteWarnings.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 处", overwriteWarnings.Count - previewCount));
            }
        }

        private static int GetCount<T>(IList<T> items)
        {
            return items == null ? 0 : items.Count;
        }
    }
}
