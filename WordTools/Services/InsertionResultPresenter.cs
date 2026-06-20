using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using WordTools.Services.Abstractions;

namespace WordTools.Services
{
    public sealed class InsertionResultPresenter
    {
        private readonly INotificationService _notificationService;
        private readonly IFailureDetailsPresenter _failureDetailsPresenter;
        private readonly IWordApplicationContext _appContext;

        public InsertionResultPresenter(
            INotificationService notificationService,
            IFailureDetailsPresenter failureDetailsPresenter,
            IWordApplicationContext appContext)
        {
            _notificationService = notificationService;
            _failureDetailsPresenter = failureDetailsPresenter;
            _appContext = appContext;
        }

        public void ShowInsertionSummary(int successCount, int failCount, string timeInfo, string timeDetail,
            List<(string fileName, string errorReason)> failedFiles,
            List<int> mergedCellRows = null,
            List<string> overwriteWarnings = null)
        {
            const int previewCount = 5;

            failedFiles = failedFiles ?? new List<(string fileName, string errorReason)>();
            mergedCellRows = mergedCellRows ?? new List<int>();
            overwriteWarnings = overwriteWarnings ?? new List<string>();

            string summaryText = InsertionSummaryFormatter.BuildSummaryMessage(
                successCount,
                failCount,
                timeInfo,
                timeDetail,
                failedFiles,
                mergedCellRows,
                overwriteWarnings,
                previewCount);

            bool showDetails = InsertionSummaryFormatter.HasMoreDetails(
                previewCount,
                failedFiles,
                mergedCellRows,
                overwriteWarnings);

            if (showDetails)
            {
                summaryText += Environment.NewLine + Environment.NewLine + InsertionSummaryFormatter.BuildDetailsPrompt();
            }

            if (showDetails)
            {
                var result = _notificationService?.ShowQuestion(
                    summaryText,
                    "插图完成",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    _failureDetailsPresenter?.ShowDetails(summaryText, failedFiles, mergedCellRows, overwriteWarnings);
                }
            }
            else
            {
                _notificationService?.ShowInformation(summaryText, "插图完成");
            }
        }

        public string BuildTimeDetail(
            bool showDetailedLog,
            long t0,
            long t1,
            long t2,
            long t3,
            long t4,
            long t5,
            bool skippedClear,
            InsertionPerformanceDiagnostics diagnostics)
        {
            if (!showDetailedLog)
            {
                return string.Empty;
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine("[诊断]");
            sb.AppendLine(string.Format("初始化: {0}ms", t0));
            sb.AppendLine(string.Format("清理编号: {0}ms (跳过={1})", t1 - t0, skippedClear));
            sb.AppendLine(string.Format("计算起始号: {0}ms", t2 - t1));
            sb.AppendLine(string.Format("预分配行: {0}ms", t3 - t2));
            sb.AppendLine(string.Format("插入图片: {0}ms", t4 - t3));
            sb.AppendLine(string.Format("收尾: {0}ms", t5 - t4));

            if (diagnostics != null)
            {
                sb.AppendLine();
                sb.Append(diagnostics.BuildDetailedLog());
            }

            return sb.ToString();
        }

        public void ShowFailureSummary(int successCount, int failCount, string timeInfo, string timeDetail,
            List<(string fileName, string errorReason)> failedFiles,
            List<int> mergedCellRows = null,
            List<string> overwriteWarnings = null)
        {
            const int previewCount = 5;

            failedFiles = failedFiles ?? new List<(string fileName, string errorReason)>();
            mergedCellRows = mergedCellRows ?? new List<int>();
            overwriteWarnings = overwriteWarnings ?? new List<string>();

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("图片插入完成：");
            sb.AppendLine(string.Format("成功: {0} 张", successCount));
            sb.AppendLine(string.Format("失败: {0} 张", failCount));
            if (mergedCellRows != null && mergedCellRows.Count > 0)
            {
                sb.AppendLine(string.Format("合并单元格绕开: {0} 处", mergedCellRows.Count));
            }
            sb.AppendLine(string.Format("耗时: {0}", timeInfo));
            sb.AppendLine();
            sb.AppendLine("失败详情（前5项）：");

            int showCount = System.Math.Min(failedFiles.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine(string.Format("  {0}: {1}", failedFiles[i].fileName, failedFiles[i].errorReason));
            }

            if (failedFiles.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 个文件失败", failedFiles.Count - previewCount));
            }

            // 合并单元格信息（默认显示 5 处）
            if (mergedCellRows != null && mergedCellRows.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("合并单元格位置（已自动绕开）：");
                int mergeShowCount = System.Math.Min(mergedCellRows.Count, previewCount);
                for (int i = 0; i < mergeShowCount; i++)
                {
                    sb.AppendLine(string.Format("  第 {0} 行", mergedCellRows[i]));
                }
                if (mergedCellRows.Count > previewCount)
                {
                    sb.AppendLine(string.Format("  ... 还有 {0} 处", mergedCellRows.Count - previewCount));
                }
            }

            sb.Append(timeDetail);

            bool hasMoreFailures = failedFiles.Count > previewCount;
            bool hasMoreMerged = mergedCellRows != null && mergedCellRows.Count > previewCount;
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            if (hasMoreFailures || hasMoreMerged)
            {
                buttons = MessageBoxButtons.YesNo;
            }

            if (buttons == MessageBoxButtons.YesNo)
            {
                var result = _notificationService?.ShowQuestion(
                    sb.ToString(),
                    "插图完成",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    _failureDetailsPresenter?.ShowDetails(sb.ToString(), failedFiles, mergedCellRows, null);
                }
            }
            else
            {
                _notificationService?.ShowInformation(sb.ToString(), "插图完成");
            }
        }

        public void ShowMergedCellWarning(List<int> mergedCellRows)
        {
            if (mergedCellRows == null || mergedCellRows.Count == 0) return;

            const int previewCount = 5;

            var sb = new System.Text.StringBuilder();
            sb.AppendLine(string.Format("检测到 {0} 处合并单元格，已自动绕开并在下方新建行插入图片。", mergedCellRows.Count));
            sb.AppendLine();
            sb.AppendLine("涉及位置（前5处）：");

            int showCount = System.Math.Min(mergedCellRows.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine(string.Format("  第 {0} 行", mergedCellRows[i]));
            }

            if (mergedCellRows.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 处", mergedCellRows.Count - previewCount));
            }

            if (mergedCellRows.Count > previewCount)
            {
                var result = _notificationService?.ShowQuestion(
                    sb.ToString(),
                    "合并单元格提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    _failureDetailsPresenter?.ShowDetails(sb.ToString(), null, mergedCellRows, null);
                }
            }
            else
            {
                _notificationService?.ShowWarning(sb.ToString(), "合并单元格提示");
            }
        }

        public void ShowOverwriteWarning(List<string> overwriteWarnings)
        {
            if (overwriteWarnings == null || overwriteWarnings.Count == 0)
            {
                return;
            }

            const int previewCount = 5;

            var sb = new System.Text.StringBuilder();
            sb.AppendLine(string.Format("检测到 {0} 处单元格已有图片或文本，已按当前规则覆盖插入新图片。", overwriteWarnings.Count));
            sb.AppendLine();
            sb.AppendLine("涉及位置（前5处）：");

            int showCount = System.Math.Min(overwriteWarnings.Count, previewCount);
            for (int i = 0; i < showCount; i++)
            {
                sb.AppendLine("  " + overwriteWarnings[i]);
            }

            if (overwriteWarnings.Count > previewCount)
            {
                sb.AppendLine(string.Format("  ... 还有 {0} 处", overwriteWarnings.Count - previewCount));
            }

            if (overwriteWarnings.Count > previewCount)
            {
                var result = _notificationService?.ShowQuestion(
                    sb.ToString(),
                    "覆盖插图提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    _failureDetailsPresenter?.ShowDetails(sb.ToString(), null, null, overwriteWarnings);
                }
            }
            else
            {
                _notificationService?.ShowWarning(sb.ToString(), "覆盖插图提示");
            }
        }

        public void TryWriteBenchmarkLog(BenchmarkLogEntry entry)
        {
            if (entry == null || !LoggingOptionsStateController.ShouldWriteBenchmarkLog(
                ConfigService.GetDetailedLoggingEnabled(),
                ConfigService.GetBenchmarkLoggingEnabled()))
            {
                return;
            }

            try
            {
                string documentPath = null;
                try
                {
                    documentPath = _appContext.Application != null && _appContext.Application.ActiveDocument != null
                        ? _appContext.Application.ActiveDocument.FullName
                        : null;
                }
                catch (Exception ex)
                {
                    SafeIgnore(ex, "获取文档路径失败");
                    documentPath = null;
                }

                string logPath = BenchmarkLogService.GetDefaultLogPath(documentPath);
                entry.DocumentPath = documentPath;
                entry.LogPath = logPath;
                BenchmarkLogService.AppendCsv(logPath, entry);
            }
            catch (Exception ex)
            {
                // 基准日志仅用于开发调试，禁止影响主流程
                SafeIgnore(ex, "写入基准日志失败");
            }
        }

        private static void SafeIgnore(Exception ex, string context)
        {
            Debug.WriteLine($"{context}: {ex.Message}");
        }
    }
}
