using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordTools.Forms;
using WordTools.Services.Abstractions;
using WordTools.Services.Adapters;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordTools.Services
{
    /// <summary>
    /// 进度服务
    /// 管理批量插图进度，提供性能优化支持
    /// </summary>
    public class ProgressService
    {
        private readonly IWordApplicationContext _appContext;
        private readonly IProgressReporter _progressReporter;
        private readonly IFailureDetailsPresenter _failureDetailsPresenter;
        private readonly INotificationService _notificationService;
        private readonly EscapeKeyMonitor _escapeMonitor;
        private readonly HighPerformanceModeController _perfController;
        private readonly InsertionResultPresenter _resultPresenter;

        // 批量处理配置
        private int _refreshInterval = 10;
        private int _memoryCleanInterval = 50;
        private int _saveInterval = 200;
        private int _fullGcInterval;  // 全代 GC 间隔
        private int _statusBarUpdateInterval = 1;  // 状态栏更新间隔（每 N 张图片更新一次）
        private DateTime _lastStatusBarUpdate = DateTime.MinValue;  // 上次状态栏更新时间
        private const int STATUS_BAR_VISIBLE_MS = 500;  // 状态栏最小可见时间（毫秒），确保用户能看清

        private InsertionPerformanceDiagnostics _activeDiagnostics;

        public ProgressService(
            IWordApplicationContext appContext,
            IProgressReporter progressReporter,
            IFailureDetailsPresenter failureDetailsPresenter,
            INotificationService notificationService)
        {
            _appContext = appContext ?? throw new ArgumentNullException(nameof(appContext));
            _progressReporter = progressReporter;
            _failureDetailsPresenter = failureDetailsPresenter;
            _notificationService = notificationService;
            _escapeMonitor = new EscapeKeyMonitor(appContext);
            _perfController = new HighPerformanceModeController(appContext);
            _resultPresenter = new InsertionResultPresenter(notificationService, failureDetailsPresenter, appContext);
        }

        #region 性能优化

        /// <summary>
        /// 清理内存（分级 GC 策略）
        /// </summary>
        private void CleanupMemory(int processedCount)
        {
            try
            {
                _appContext.DoEvents();

                // 内存水位检查：超过 500MB 时强制全代回收
                if (GC.GetTotalMemory(false) > 500 * 1024 * 1024)
                {
                    GC.Collect(2);
                    GC.WaitForPendingFinalizers();
                }
                else if (processedCount % _fullGcInterval == 0)
                {
                    // 定期全代回收（防止大对象堆膨胀）
                    GC.Collect(2);
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    // 常规第 0 代快速回收
                    GC.Collect(0);
                }
            }
            catch (Exception ex)
            {
                SafeIgnore(ex, "清理内存失败");
            }
        }

        #endregion

        #region 批量插入 - 从文件夹

        /// <summary>
        /// 批量插图主入口（带进度）
        /// </summary>
        public void InsertPhotosWithProgress(string folderPath, float minHeight,
            bool needDescription, bool useFileNameAsDescription, bool useFolderNameAsDescription,
            bool includeRootImages, bool includeSubFolderImages,
            bool needAutoNumbering, int numberAlignment = 2, int numberPosition = 1)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            long t0 = 0, t1 = 0, t2 = 0, t3 = 0, t4 = 0, t5 = 0;
            bool skippedClear = false;
            string benchmarkStatus = "Completed";
            string benchmarkError = null;
            _escapeMonitor.Reset();
            int processedCount = 0;
            int successCount = 0;
            int failCount = 0;
            int totalFiles = 0;
            var failedFiles = new List<(string fileName, string errorReason)>();
            var mergedCellRows = new List<int>();
            var overwriteWarnings = new List<string>();
            var insertionDiagnostics = new InsertionPerformanceDiagnostics();
            _activeDiagnostics = insertionDiagnostics;
            var batchContext = new ImageInsertionBatchContext(insertionDiagnostics);
            DateTime startTime = DateTime.Now;
            Table tbl = null;
            int startRow = 0;

            try
            {
                // 验证表格
                var selection = _appContext.Application.Selection;
                if (!TableService.IsSelectionInTable(selection))
                {
                    _notificationService?.ShowWarning("请先选中一个表格！", "提示");
                    return;
                }

                if (!TableService.IsSelectionInFirstColumn(selection))
                {
                    _notificationService?.ShowWarning("请将光标置于表格左侧单元格！", "提示");
                    return;
                }

                tbl = TableService.GetCurrentTable(selection);
                TableService.AdjustTableColumns(tbl, 2);
                int rowIndex = selection.Cells[1].RowIndex;
                startRow = rowIndex;

                // 设置固定列宽
                if (!TableService.IsTableFixedColumnWidth(tbl))
                {
                    TableService.SetTableFixedColumnWidth(tbl);
                }

                t0 = sw.ElapsedMilliseconds;

                // 清除 startRow 之后的编号（增量模式，不影响前面已有的编号）
                // 只在 startRow 位于现有内容范围内且后续存在实际内容时才清理
                int currentRowCount = tbl.Rows.Count;
                bool needClearNumbering = false;
                if (startRow <= currentRowCount)
                {
                    // 轻量检查：仅检查起始行第一列的文本内容
                    try
                    {
                        Range checkRange = tbl.Cell(startRow, 1).Range;
                        checkRange.SetRange(checkRange.Start, checkRange.End - 1);
                        string text = (checkRange.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        // 检查文本是否已有编号或包含图片标记
                        bool hasContent = !string.IsNullOrEmpty(text);
                        bool hasPictureMarker = text.IndexOf('\x01') >= 0; // 图片在 Range.Text 中表示为 \x01
                        
                        // 只有起始行已有内容时才需要清理编号
                        if (hasContent || hasPictureMarker)
                        {
                            needClearNumbering = true;
                        }
                    }
                    catch (Exception ex) { SafeIgnore(ex, "检测起始行内容失败，保守清理编号"); needClearNumbering = true; } // 出错时保守处理，执行清理
                }
                
                if (needClearNumbering)
                {
                    TableNumberingService.ClearTableNumbering(tbl, startRow);
                }
                else
                {
                    skippedClear = true;
                }
                t1 = sw.ElapsedMilliseconds;

                // 一次遍历获取图片文件列表和总数，避免重复扫描目录
                var imageFiles = FileService.GetImageFiles(folderPath, includeRootImages, includeSubFolderImages);
                totalFiles = imageFiles.TotalCount;

                if (totalFiles == 0)
                {
                    _notificationService?.ShowWarning("未找到任何图片文件！", "提示");
                    return;
                }

                // 仅在图片数量较多时显示开始提示，并让用户有机会取消
                if (totalFiles > 20)
                {
                    var confirmResult = _notificationService?.ShowQuestion(
                        string.Format("开始插入图片...\n\n提示：插入过程中可以按 ESC 键随时取消操作。\n\n共找到 {0} 张图片。", totalFiles),
                        "批量插图",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Information) ?? DialogResult.OK;

                    if (confirmResult != DialogResult.OK)
                    {
                        _escapeMonitor.Cancel();
                        return;
                    }
                }

                // 设置优化参数
                _refreshInterval = _perfController.GetOptimizedRefreshInterval(totalFiles);
                _statusBarUpdateInterval = _perfController.GetStatusBarUpdateInterval(totalFiles);
                _memoryCleanInterval = _refreshInterval * 10;
                _fullGcInterval = _memoryCleanInterval * 5;
                _saveInterval = _refreshInterval * 20;

                // 进入高性能模式
                _perfController.Enter();

                // 创建并显示进度窗口
                _progressReporter?.Show();
                _appContext.DoEvents();

                // 编号准备
                WdParagraphAlignment wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter;
                switch (numberAlignment)
                {
                    case 1: wdAlignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                    case 3: wdAlignment = WdParagraphAlignment.wdAlignParagraphRight; break;
                    default: wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter; break;
                }
                bool isFirstSeqField = true;
                int startNumber = 1;
                if (needAutoNumbering && startRow > 1)
                {
                    startNumber = TableNumberingService.CalculateNextSequenceNumber(tbl, startRow);
                }
                t2 = sw.ElapsedMilliseconds;
                int currentNumber = startNumber;

                // 预分配行数
                ImageService.PreAllocateRows(tbl, totalFiles, 2, needDescription, _appContext.Application);
                t3 = sw.ElapsedMilliseconds;

                // 进度窗口已显示，无需额外状态栏提示
                // 获取文件夹名称
                string rootFolderName = FileService.GetFolderName(folderPath);

                // 处理根目录图片
                if (includeRootImages && imageFiles.RootFiles != null && imageFiles.RootFiles.Length > 0)
                {
                    // 创建标题行
                    TableService.CreateTitleRow(tbl, ref rowIndex, rootFolderName);
                    batchContext.ClearRowAvailability();

                    // 处理文件
                    ProcessFileBatch(imageFiles.RootFiles, tbl, ref rowIndex, minHeight, needDescription,
                        useFileNameAsDescription, useFolderNameAsDescription, rootFolderName, ref processedCount, ref successCount, ref failCount,
                        totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber, failedFiles, mergedCellRows, overwriteWarnings, batchContext);
                }

                // 处理子文件夹
                if (includeSubFolderImages && imageFiles.SubfolderFiles != null)
                {
                    var subfolderKeys = new System.Collections.Generic.List<string>(imageFiles.SubfolderFiles.Keys);
                    foreach (var subfolder in subfolderKeys)
                    {
                        if (_escapeMonitor.ShouldCancel(_progressReporter)) break;

                        string[] subFiles = imageFiles.SubfolderFiles[subfolder];

                        if (subFiles != null && subFiles.Length > 0)
                        {
                            string subfolderName = FileService.GetFolderName(subfolder);
                            TableService.CreateTitleRow(tbl, ref rowIndex, subfolderName);
                            batchContext.ClearRowAvailability();

                            ProcessFileBatch(subFiles, tbl, ref rowIndex, minHeight, needDescription,
                                useFileNameAsDescription, useFolderNameAsDescription, subfolderName, ref processedCount, ref successCount, ref failCount,
                                totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber, failedFiles, mergedCellRows, overwriteWarnings, batchContext);

                            // 释放引用，帮助垃圾回收
                            imageFiles.SubfolderFiles[subfolder] = null;
                        }
                    }
                }

                // 记录 t4（图片插入完成时间）
                t4 = sw.ElapsedMilliseconds;

                // 显示完成消息
                stopwatch.Stop();
                double seconds = stopwatch.Elapsed.TotalSeconds;
                string timeInfo = seconds >= 60
                    ? string.Format("{0}分钟{1:F1}秒", (int)(seconds / 60), seconds % 60)
                    : string.Format("{0:F1}秒", seconds);

                // 记录 t5（最终时间），用于计算收尾耗时
                t5 = sw.ElapsedMilliseconds;
                bool showDetailedLog = LoggingOptionsStateController.ShouldShowDetailedLog(
                    ConfigService.GetDetailedLoggingEnabled());

                string timeDetail = _resultPresenter.BuildTimeDetail(
                    showDetailedLog,
                    t0,
                    t1,
                    t2,
                    t3,
                    t4,
                    t5,
                    skippedClear,
                    insertionDiagnostics);

                // 先将进度窗口标记为完成状态，避免关闭时触发"是否取消"提示
                _progressReporter?.ShowCompletion(successCount, failCount, seconds);
                // 更新进度窗口为完成状态
                CloseProgressForm();

                if (_escapeMonitor.IsCancelled)
                {
                    benchmarkStatus = "Cancelled";
                    _notificationService?.ShowWarning(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo) + timeDetail, "提示");
                }
                else if (InsertionSummaryFormatter.ShouldShowSummary(failCount, mergedCellRows, overwriteWarnings))
                {
                    benchmarkStatus = failCount > 0 ? "CompletedWithFailures" : "CompletedWithWarnings";
                    _resultPresenter.ShowInsertionSummary(successCount, failCount, timeInfo, timeDetail, failedFiles, mergedCellRows, overwriteWarnings);
                }
                else
                {
                    _notificationService?.ShowInformation(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo) + timeDetail, "完成");
                }
            }
            catch (Exception ex)
            {
                benchmarkStatus = "Error";
                benchmarkError = ex.Message;
                _notificationService?.ShowError(string.Format("处理过程中发生错误: {0}", ex.Message), "错误");
            }
            finally
            {
                if (_escapeMonitor.IsCancelled)
                {
                    benchmarkStatus = "Cancelled";
                }

                if (stopwatch.IsRunning)
                {
                    stopwatch.Stop();
                }

                if (t4 == 0)
                {
                    t4 = sw.ElapsedMilliseconds;
                }

                if (t5 == 0)
                {
                    t5 = sw.ElapsedMilliseconds;
                }

                // 1. 先退出高性能模式，让用户立即看到已插入的图片
                _perfController.Exit();

                // 循环 DoEvents 让 Word 有充足时间完成屏幕重绘
                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        _appContext.DoEvents();
                        System.Threading.Thread.Sleep(10);
                    }
                }
                catch (Exception ex) { SafeIgnore(ex, "等待 Word 完成消息处理失败"); }

                // 3. 确保域代码不可见（保持原有逻辑不变）
                try
                {
                    if (_appContext.Application.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
                    {
                        _appContext.Application.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        _appContext.DoEvents();
                    }
                }
                catch (Exception ex) { SafeIgnore(ex, "隐藏域代码失败"); }

                // 关闭进度窗口
                CloseProgressForm();

                _appContext.Application.StatusBar = "";

                _resultPresenter.TryWriteBenchmarkLog(new BenchmarkLogEntry
                {
                    RunMode = "Folder",
                    Status = benchmarkStatus,
                    SourcePath = folderPath,
                    TotalFiles = totalFiles,
                    ProcessedCount = processedCount,
                    SuccessCount = successCount,
                    FailCount = failCount,
                    MergedCellCount = mergedCellRows.Count,
                    Cancelled = _escapeMonitor.IsCancelled,
                    NeedDescription = needDescription,
                    UseFileNameAsDescription = useFileNameAsDescription,
                    UseFolderNameAsDescription = useFolderNameAsDescription,
                    AutoNumbering = needAutoNumbering,
                    NumberAlignment = numberAlignment,
                    NumberPosition = numberPosition,
                    TotalSeconds = stopwatch.Elapsed.TotalSeconds,
                    InitializeMs = t0,
                    ClearNumberingMs = t1 >= t0 ? t1 - t0 : (long?)null,
                    CalculateStartNumberMs = t2 >= t1 ? t2 - t1 : (long?)null,
                    PreAllocateRowsMs = t3 >= t2 ? t3 - t2 : (long?)null,
                    InsertImagesMs = t4 >= t3 ? t4 - t3 : (long?)null,
                    WrapUpMs = t5 >= t4 ? t5 - t4 : (long?)null,
                    CellAvailabilityMs = insertionDiagnostics.CellAvailabilityMs,
                    CellAvailabilityCount = insertionDiagnostics.CellAvailabilityCount,
                    FloatingShapeLookupMs = insertionDiagnostics.FloatingShapeLookupMs,
                    FloatingShapeLookupCount = insertionDiagnostics.FloatingShapeLookupCount,
                    OverwriteClearMs = insertionDiagnostics.OverwriteClearMs,
                    OverwriteClearCount = insertionDiagnostics.OverwriteClearCount,
                    AddPictureMs = insertionDiagnostics.AddPictureMs,
                    AddPictureCount = insertionDiagnostics.AddPictureCount,
                    CellValidationMs = insertionDiagnostics.CellValidationMs,
                    CellValidationCount = insertionDiagnostics.CellValidationCount,
                    PictureSizingMs = insertionDiagnostics.PictureSizingMs,
                    PictureSizingCount = insertionDiagnostics.PictureSizingCount,
                    ProgressUiMs = insertionDiagnostics.ProgressUiMs,
                    ProgressUiCount = insertionDiagnostics.ProgressUiCount,
                    DescriptionWriteMs = insertionDiagnostics.DescriptionWriteMs,
                    DescriptionWriteCount = insertionDiagnostics.DescriptionWriteCount,
                    SkippedClearNumbering = skippedClear,
                    ErrorMessage = benchmarkError
                });

                _activeDiagnostics = null;
            }
        }

        #endregion

        #region 批量插入 - 选中文件

        /// <summary>
        /// 插入选中的图片（带进度）
        /// </summary>
        public void InsertSelectedPhotosWithProgress(string[] files, float minHeight,
            bool needDescription, bool useFileNameAsDescription, bool useFolderNameAsDescription,
            bool needAutoNumbering, int numberAlignment = 2, int numberPosition = 1)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            long t0 = 0, t1 = 0, t2 = 0, t3 = 0, t4 = 0, t5 = 0;
            bool skippedClear = false;
            string benchmarkStatus = "Completed";
            string benchmarkError = null;
            _escapeMonitor.Reset();
            int processedCount = 0;
            int successCount = 0;
            int failCount = 0;
            var failedFiles = new List<(string fileName, string errorReason)>();
            var mergedCellRows = new List<int>();
            var overwriteWarnings = new List<string>();
            var insertionDiagnostics = new InsertionPerformanceDiagnostics();
            _activeDiagnostics = insertionDiagnostics;
            var batchContext = new ImageInsertionBatchContext(insertionDiagnostics);
            DateTime startTime = DateTime.Now;
            Table tbl = null;
            int startRow = 0;
            string sourcePath = files != null && files.Length > 0
                ? FileService.GetParentFolder(files[0])
                : "";

            try
            {
                // 验证表格
                var selection = _appContext.Application.Selection;
                if (!TableService.IsSelectionInTable(selection))
                {
                    _notificationService?.ShowWarning("请先选中一个表格！", "提示");
                    return;
                }

                if (!TableService.IsSelectionInFirstColumn(selection))
                {
                    _notificationService?.ShowWarning("请将光标置于表格左侧单元格！", "提示");
                    return;
                }

                tbl = TableService.GetCurrentTable(selection);
                if (!TableService.IsTableFixedColumnWidth(tbl))
                {
                    TableService.SetTableFixedColumnWidth(tbl);
                }

                TableService.AdjustTableColumns(tbl, 2);
                int rowIndex = selection.Cells[1].RowIndex;
                startRow = rowIndex;
                t0 = sw.ElapsedMilliseconds;

                int totalFiles = files.Length;
                _refreshInterval = _perfController.GetOptimizedRefreshInterval(totalFiles);
                _statusBarUpdateInterval = _perfController.GetStatusBarUpdateInterval(totalFiles);
                // 注意：ProcessFileBatch 内部会做批量预检，totalFiles 保持原始值用于进度显示

                // 清除 startRow 之后的编号（增量模式，不影响前面已有的编号）
                // 优化：只在 startRow 位于现有内容范围内且 startRow 之后有实际内容时才清理
                // 如果 startRow 已经超出当前表格行数，说明是在末尾追加，新行为空无需清理
                int currentRowCount = tbl.Rows.Count;
                bool needClearNumbering = false;
                if (startRow <= currentRowCount)
                {
                    // 超轻量检查：只检查 startRow 行第 1 列的文本内容（不含 InlineShapes.Count）
                    try
                    {
                        Range checkRange = tbl.Cell(startRow, 1).Range;
                        checkRange.SetRange(checkRange.Start, checkRange.End - 1);
                        string text = (checkRange.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        // 检查文本是否以数字开头（已有编号）或包含图片标记
                        bool hasContent = !string.IsNullOrEmpty(text);
                        bool hasPictureMarker = text.IndexOf('\x01') >= 0; // 图片在 Range.Text 中表示为 \x01
                        
                        // 如果 startRow 有内容（文本或图片标记），才需要清理编号
                        if (hasContent || hasPictureMarker)
                        {
                            needClearNumbering = true;
                        }
                        // 注意：不再遍历 tbl.Range.InlineShapes 检查后续行，该操作在大表格中极慢（5-6 秒）
                        // 如果用户从中间插入，需要清理编号，可通过其他方式检测（如检查描述行是否有编号文本）
                    }
                    catch (Exception ex) { SafeIgnore(ex, "检测起始行内容失败，保守清理编号"); needClearNumbering = true; } // 出错时保守处理，执行清理
                }
                
                if (needClearNumbering)
                {
                    TableNumberingService.ClearTableNumbering(tbl, startRow);
                }
                else
                {
                    skippedClear = true;
                }
                t1 = sw.ElapsedMilliseconds;

                _perfController.Enter();

                // 创建并显示进度窗口（独立于 Word，不受 ScreenUpdating 影响）
                _progressReporter?.Show();
                _appContext.DoEvents();

                // 编号准备
                WdParagraphAlignment wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter;
                switch (numberAlignment)
                {
                    case 1: wdAlignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                    case 3: wdAlignment = WdParagraphAlignment.wdAlignParagraphRight; break;
                    default: wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter; break;
                }
                bool isFirstSeqField = true;
                int startNumber = 1;
                if (needAutoNumbering && startRow > 1)
                {
                    startNumber = TableNumberingService.CalculateNextSequenceNumber(tbl, startRow);
                }
                t2 = sw.ElapsedMilliseconds;
                t3 = t2;
                int currentNumber = startNumber;

                // 处理文件
                string folderName = FileService.GetFolderName(FileService.GetParentFolder(files[0]));
                ProcessFileBatch(files, tbl, ref rowIndex, minHeight, needDescription,
                    useFileNameAsDescription, useFolderNameAsDescription, folderName, ref processedCount, ref successCount, ref failCount,
                    totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber, failedFiles, mergedCellRows, overwriteWarnings, batchContext);
                t4 = sw.ElapsedMilliseconds;

                stopwatch.Stop();
                double seconds = stopwatch.Elapsed.TotalSeconds;
                string timeInfo = seconds >= 60
                    ? string.Format("{0}分钟{1:F1}秒", (int)(seconds / 60), seconds % 60)
                    : string.Format("{0:F1}秒", seconds);
                t5 = sw.ElapsedMilliseconds;
                bool showDetailedLog = LoggingOptionsStateController.ShouldShowDetailedLog(
                    ConfigService.GetDetailedLoggingEnabled());
                string timeDetail = _resultPresenter.BuildTimeDetail(
                    showDetailedLog,
                    t0,
                    t1,
                    t2,
                    t3,
                    t4,
                    t5,
                    skippedClear,
                    insertionDiagnostics);

                // 先将进度窗口标记为完成状态，避免关闭时触发“是否取消”提示
                _progressReporter?.ShowCompletion(successCount, failCount, seconds);
                // 更新进度窗口为完成状态
                CloseProgressForm();

                if (_escapeMonitor.IsCancelled)
                {
                    benchmarkStatus = "Cancelled";
                    _notificationService?.ShowWarning(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo) + timeDetail, "提示");
                }
                else if (InsertionSummaryFormatter.ShouldShowSummary(failCount, mergedCellRows, overwriteWarnings))
                {
                    benchmarkStatus = failCount > 0 ? "CompletedWithFailures" : "CompletedWithWarnings";
                    _resultPresenter.ShowInsertionSummary(successCount, failCount, timeInfo, timeDetail, failedFiles, mergedCellRows, overwriteWarnings);
                }
                else
                {
                    _notificationService?.ShowInformation(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo) + timeDetail, "完成");
                }
            }
            catch (Exception ex)
            {
                benchmarkStatus = "Error";
                benchmarkError = ex.Message;
                _notificationService?.ShowError(string.Format("处理过程中发生错误: {0}", ex.Message), "错误");
            }
            finally
            {
                if (_escapeMonitor.IsCancelled)
                {
                    benchmarkStatus = "Cancelled";
                }

                if (stopwatch.IsRunning)
                {
                    stopwatch.Stop();
                }

                if (t4 == 0)
                {
                    t4 = sw.ElapsedMilliseconds;
                }

                if (t5 == 0)
                {
                    t5 = sw.ElapsedMilliseconds;
                }

                // 1. 先退出高性能模式，让用户立即看到已插入的图片
                _perfController.Exit();

                // 循环 DoEvents 让 Word 有充足时间完成屏幕重绘
                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        _appContext.DoEvents();
                        System.Threading.Thread.Sleep(10);
                    }
                }
                catch (Exception ex) { SafeIgnore(ex, "等待 Word 完成消息处理失败"); }

                // 3. 确保域代码不可见（保持原有逻辑不变）
                try
                {
                    if (_appContext.Application.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
                    {
                        _appContext.Application.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        _appContext.DoEvents();
                    }
                }
                catch (Exception ex) { SafeIgnore(ex, "隐藏域代码失败"); }

                // 关闭进度窗口
                CloseProgressForm();

                _appContext.Application.StatusBar = "";

                _resultPresenter.TryWriteBenchmarkLog(new BenchmarkLogEntry
                {
                    RunMode = "SelectedFiles",
                    Status = benchmarkStatus,
                    SourcePath = sourcePath,
                    TotalFiles = files != null ? files.Length : 0,
                    ProcessedCount = processedCount,
                    SuccessCount = successCount,
                    FailCount = failCount,
                    MergedCellCount = mergedCellRows.Count,
                    Cancelled = _escapeMonitor.IsCancelled,
                    NeedDescription = needDescription,
                    UseFileNameAsDescription = useFileNameAsDescription,
                    UseFolderNameAsDescription = useFolderNameAsDescription,
                    AutoNumbering = needAutoNumbering,
                    NumberAlignment = numberAlignment,
                    NumberPosition = numberPosition,
                    TotalSeconds = stopwatch.Elapsed.TotalSeconds,
                    InitializeMs = t0,
                    ClearNumberingMs = t1 >= t0 ? t1 - t0 : (long?)null,
                    CalculateStartNumberMs = t2 >= t1 ? t2 - t1 : (long?)null,
                    InsertImagesMs = t4 >= t2 ? t4 - t2 : (long?)null,
                    WrapUpMs = t5 >= t4 ? t5 - t4 : (long?)null,
                    CellAvailabilityMs = insertionDiagnostics.CellAvailabilityMs,
                    CellAvailabilityCount = insertionDiagnostics.CellAvailabilityCount,
                    FloatingShapeLookupMs = insertionDiagnostics.FloatingShapeLookupMs,
                    FloatingShapeLookupCount = insertionDiagnostics.FloatingShapeLookupCount,
                    OverwriteClearMs = insertionDiagnostics.OverwriteClearMs,
                    OverwriteClearCount = insertionDiagnostics.OverwriteClearCount,
                    AddPictureMs = insertionDiagnostics.AddPictureMs,
                    AddPictureCount = insertionDiagnostics.AddPictureCount,
                    CellValidationMs = insertionDiagnostics.CellValidationMs,
                    CellValidationCount = insertionDiagnostics.CellValidationCount,
                    PictureSizingMs = insertionDiagnostics.PictureSizingMs,
                    PictureSizingCount = insertionDiagnostics.PictureSizingCount,
                    ProgressUiMs = insertionDiagnostics.ProgressUiMs,
                    ProgressUiCount = insertionDiagnostics.ProgressUiCount,
                    DescriptionWriteMs = insertionDiagnostics.DescriptionWriteMs,
                    DescriptionWriteCount = insertionDiagnostics.DescriptionWriteCount,
                    SkippedClearNumbering = skippedClear,
                    ErrorMessage = benchmarkError
                });

                _activeDiagnostics = null;
            }
        }

        #endregion

        #region 文件批量处理

        /// <summary>
        /// 处理一批文件
        /// </summary>
        private void ProcessFileBatch(string[] files, Table tbl, ref int rowIndex,
            float minHeight, bool needDescription, bool useFileNameAsDescription, bool useFolderNameAsDescription,
            string folderName,
            ref int processedCount, ref int successCount, ref int failCount,
            int totalFiles, DateTime startTime,
            bool needAutoNumbering, WdParagraphAlignment numberAlignment, int numberPosition,
            ref bool isFirstSeqField, ref int currentNumber,
            List<(string fileName, string errorReason)> failedFiles = null,
            List<int> mergedCellRows = null,
            List<string> overwriteWarnings = null,
            ImageInsertionBatchContext batchContext = null)
        {
            var currentRowFiles = new List<string>();

            // 性能优化：缓存行数，避免重复 COM 调用
            int cachedRowCount = tbl.Rows.Count;
            int imagesPlacedInCurrentRow = 0;

            // 性能优化：循环前预先调整列数
            TableService.AdjustTableColumns(tbl, 2);

            // 性能优化：批量预检文件，避免循环内重复磁盘 IO
            var validFiles = FileService.BatchValidateImageFiles(files, failedFiles);
            int precheckFailedCount = files.Length - validFiles.Count;
            failCount += precheckFailedCount;

            for (int i = 0; i < validFiles.Count; i++)
            {
                if (_escapeMonitor.ShouldCancel(_progressReporter)) break;

                string filePath = validFiles[i];
                string fileName = FileService.GetFileName(filePath);

                // 更新进度窗口
                if (processedCount % _statusBarUpdateInterval == 0)
                {
                    UpdateStatusBar(processedCount, totalFiles, fileName, startTime);
                    // 处理消息队列，确保进度窗口 UI 更新
                    _appContext.DoEvents();
                }

                try
                {
                    // 确保行存在（使用缓存行数优化）
                    if (rowIndex > cachedRowCount)
                    {
                        int rowCountBefore = cachedRowCount;
                        tbl.Rows.Add();
                        cachedRowCount = tbl.Rows.Count;
                        if (batchContext != null)
                        {
                            batchContext.ClearRowAvailability();
                        }
                        if (cachedRowCount <= rowCountBefore)
                        {
                            throw new InvalidOperationException("表格行添加失败");
                        }
                    }

                    int colIndex = imagesPlacedInCurrentRow == 0 ? 1 : 2;
                    ImageCellAvailability targetCellAvailability = ImageCellAvailability.Available;

                    // 列数保护：确保目标列存在
                    if (tbl.Columns.Count < colIndex)
                    {
                        TableService.AdjustTableColumns(tbl, 2);
                    }

                    if (imagesPlacedInCurrentRow == 0)
                    {
                        int foundRow;
                        if (TableService.FindNextSuitableImageRow(tbl, rowIndex, out foundRow, mergedCellRows, batchContext))
                        {
                            rowIndex = foundRow;
                        }
                        else
                        {
                            rowIndex = tbl.Rows.Count + 1;
                            tbl.Rows.Add();
                            cachedRowCount = tbl.Rows.Count;
                            TableService.AdjustTableColumns(tbl, 2);
                            if (batchContext != null)
                            {
                                batchContext.ClearRowAvailability();
                            }
                        }
                        colIndex = 1;
                        targetCellAvailability = TableService.GetCellAvailability(tbl, rowIndex, colIndex, batchContext);
                    }
                    else
                    {
                        var secondCellAvailability = TableService.GetCellAvailability(tbl, rowIndex, 2, batchContext);
                        if (secondCellAvailability == ImageCellAvailability.Merged)
                        {
                            AddMergedRow(mergedCellRows, rowIndex);
                        }

                        if (!ImageRowPlanner.CanHostSingleImage(secondCellAvailability))
                        {
                            imagesPlacedInCurrentRow = 0;
                            currentRowFiles.Clear();
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                            if (batchContext != null)
                            {
                                batchContext.ClearRowAvailability();
                            }

                            if (ImageRowPlanner.ShouldRetryCurrentImage(secondCellAvailability))
                            {
                                i--;
                                continue;
                            }

                            failCount++;
                            failedFiles?.Add((fileName, "当前行第二列不可用，已跳过该行以保持每行两张图片"));
                            processedCount++;
                            if (processedCount % _memoryCleanInterval == 0)
                            {
                                CleanupMemory(processedCount);
                            }
                            continue;
                        }

                        targetCellAvailability = secondCellAvailability;
                    }

                    // 插入图片
                    string errorMsg;
                    bool inserted = ImageService.InsertImageFast(tbl.Cell(rowIndex, colIndex), filePath, out errorMsg, minHeight, batchContext);
                    if (inserted)
                    {
                        if (batchContext != null)
                        {
                            batchContext.ClearRowAvailability();
                        }

                        if (ImageRowPlanner.RequiresOverwriteWarning(targetCellAvailability))
                        {
                            AddOverwriteWarning(overwriteWarnings, rowIndex, colIndex, targetCellAvailability);
                        }

                        successCount++;

                        if (useFileNameAsDescription)
                        {
                            currentRowFiles.Add(filePath);
                        }
                        else if (useFolderNameAsDescription)
                        {
                            string cleanedFolderName = System.Text.RegularExpressions.Regex.Replace(folderName, @"^\d+\s*[\.\-]?\s*", "");
                            currentRowFiles.Add(cleanedFolderName);
                        }

                        imagesPlacedInCurrentRow++;

                        if (imagesPlacedInCurrentRow == 2)
                        {
                            if (useFileNameAsDescription || useFolderNameAsDescription)
                            {
                                var descriptionWriteWatch = System.Diagnostics.Stopwatch.StartNew();
                                rowIndex++;
                                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                                int descriptionRow = rowIndex;
                                bool descriptionsAreFilePaths = useFileNameAsDescription && !useFolderNameAsDescription;
                                TableService.InsertFileNameDescriptionRow(tbl, ref rowIndex, currentRowFiles.ToArray(), descriptionsAreFilePaths);
                                if (needAutoNumbering)
                                {
                                    for (int col = 1; col <= 2; col++)
                                    {
                                        try { TableNumberingService.InsertNumberText(tbl, descriptionRow, col, numberAlignment, currentNumber, true, numberPosition); }
                                        catch (Exception ex) { SafeIgnore(ex, "插入编号文本失败"); }
                                        currentNumber++;
                                    }
                                }
                                currentRowFiles.Clear();
                                descriptionWriteWatch.Stop();
                                if (batchContext != null)
                                {
                                    batchContext.Diagnostics.RecordDescriptionWrite(descriptionWriteWatch.ElapsedMilliseconds);
                                    batchContext.ClearRowAvailability();
                                }
                            }
                            else if (needDescription)
                            {
                                var descriptionWriteWatch = System.Diagnostics.Stopwatch.StartNew();
                                rowIndex++;
                                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                                TableService.InsertDescriptionRow(tbl, ref rowIndex);
                                if (needAutoNumbering)
                                {
                                    for (int col = 1; col <= 2; col++)
                                    {
                                        try { TableNumberingService.InsertNumberText(tbl, rowIndex, col, numberAlignment, currentNumber, true, numberPosition); }
                                        catch (Exception ex) { SafeIgnore(ex, "插入编号文本失败"); }
                                        currentNumber++;
                                    }
                                }
                                rowIndex++;
                                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                                descriptionWriteWatch.Stop();
                                if (batchContext != null)
                                {
                                    batchContext.Diagnostics.RecordDescriptionWrite(descriptionWriteWatch.ElapsedMilliseconds);
                                    batchContext.ClearRowAvailability();
                                }
                            }
                            else
                            {
                                rowIndex++;
                                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                                if (batchContext != null)
                                {
                                    batchContext.ClearRowAvailability();
                                }
                            }

                            imagesPlacedInCurrentRow = 0;
                        }
                    }
                    else
                    {
                        failCount++;
                        failedFiles?.Add((fileName, errorMsg ?? "未知错误"));

                        if (imagesPlacedInCurrentRow == 1 && InsertionErrorClassifier.IsMergedCellError(errorMsg))
                        {
                            AddMergedRow(mergedCellRows, rowIndex);
                            imagesPlacedInCurrentRow = 0;
                            currentRowFiles.Clear();
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                            if (batchContext != null)
                            {
                                batchContext.ClearRowAvailability();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    string classifiedError = InsertionErrorClassifier.Classify(ex);
                    failedFiles?.Add((fileName, classifiedError));

                    if (imagesPlacedInCurrentRow == 1 && InsertionErrorClassifier.IsMergedCellError(classifiedError))
                    {
                        AddMergedRow(mergedCellRows, rowIndex);
                        imagesPlacedInCurrentRow = 0;
                        currentRowFiles.Clear();
                        rowIndex++;
                        TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                        if (batchContext != null)
                        {
                            batchContext.ClearRowAvailability();
                        }
                    }
                }

                processedCount++;

                // 定期清理内存
                if (processedCount % _memoryCleanInterval == 0)
                {
                    CleanupMemory(processedCount);
                }
            }

            // 处理最后一行（如果最后一行仅有 1 张图片，需要填充空单元格）
            if (imagesPlacedInCurrentRow == 1)
            {
                TableService.FillEmptyCellsWithNA(tbl, rowIndex, 2, 2);
            }

            // 处理末尾的描述行
            if ((useFileNameAsDescription || useFolderNameAsDescription) && currentRowFiles.Count > 0)
            {
                var descriptionWriteWatch = System.Diagnostics.Stopwatch.StartNew();
                rowIndex++;
                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                // 保存当前行号，因为 InsertFileNameDescriptionRow 会递增 rowIndex
                int descriptionRow = rowIndex;
                bool descriptionsAreFilePaths = useFileNameAsDescription && !useFolderNameAsDescription;
                TableService.InsertFileNameDescriptionRow(tbl, ref rowIndex, currentRowFiles.ToArray(), descriptionsAreFilePaths);
                // 内联编号，每列独立递增
                if (needAutoNumbering)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        try { TableNumberingService.InsertNumberText(tbl, descriptionRow, col, numberAlignment, currentNumber, true, numberPosition); }
                        catch (Exception ex) { SafeIgnore(ex, "插入编号文本失败"); }
                        currentNumber++; // 每列递增
                    }
                }
                descriptionWriteWatch.Stop();
                if (batchContext != null)
                {
                    batchContext.Diagnostics.RecordDescriptionWrite(descriptionWriteWatch.ElapsedMilliseconds);
                    batchContext.ClearRowAvailability();
                }
            }
            else if (needDescription && imagesPlacedInCurrentRow > 0)
            {
                var descriptionWriteWatch = System.Diagnostics.Stopwatch.StartNew();
                rowIndex++;
                TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                TableService.InsertDescriptionRow(tbl, ref rowIndex);
                // 内联编号，每列独立递增
                if (needAutoNumbering)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        try { TableNumberingService.InsertNumberText(tbl, rowIndex, col, numberAlignment, currentNumber, true, numberPosition); }
                        catch (Exception ex) { SafeIgnore(ex, "插入编号文本失败"); }
                        currentNumber++; // 每列递增
                    }
                }
                rowIndex++;
                descriptionWriteWatch.Stop();
                if (batchContext != null)
                {
                    batchContext.Diagnostics.RecordDescriptionWrite(descriptionWriteWatch.ElapsedMilliseconds);
                    batchContext.ClearRowAvailability();
                }
            }

            // 循环结束后统一调整列数（性能优化：避免每次 EnsureRowExists 都调用）
            TableService.AdjustTableColumns(tbl, 2);
        }

        #endregion

        #region 状态栏更新

        /// <summary>
        /// 更新进度窗口和状态栏
        /// </summary>
        private void UpdateStatusBar(int current, int total, string currentFile, DateTime startTime)
        {
            var uiWatch = _activeDiagnostics != null
                ? Stopwatch.StartNew()
                : null;

            try
            {
                // 检查距离上次更新是否超过最小可见时间，避免更新过于频繁
                var now = DateTime.Now;
                var elapsedSinceLastUpdate = (now - _lastStatusBarUpdate).TotalMilliseconds;
                if (elapsedSinceLastUpdate < STATUS_BAR_VISIBLE_MS && _lastStatusBarUpdate != DateTime.MinValue)
                {
                    return;
                }
                _lastStatusBarUpdate = now;

                int percent = total > 0 ? (int)((double)current / total * 100) : 0;
                var elapsed = now - startTime;

                string shortFileName = currentFile.Length > 30
                    ? "..." + currentFile.Substring(currentFile.Length - 27)
                    : currentFile;

                // 更新进度窗口（独立于 Word，不受 ScreenUpdating 影响）
                _progressReporter?.UpdateProgress(current, total, shortFileName, elapsed);
                // 确保进度窗口保持置顶
                WindowActivationService.EnsureWindowTopMost(_progressReporter?.Handle ?? IntPtr.Zero);

                // 进度由独立窗口显示，不再更新 Word 状态栏
            }
            catch (Exception ex)
            {
                SafeIgnore(ex, "更新进度窗口失败");
            }
            finally
            {
                if (uiWatch != null)
                {
                    uiWatch.Stop();
                    _activeDiagnostics.RecordProgressUi(uiWatch.ElapsedMilliseconds);
                }
            }
        }

        #endregion

        #region 错误分类

        private static void AddMergedRow(List<int> mergedCellRows, int rowIndex)
        {
            if (mergedCellRows == null || rowIndex < 1)
            {
                return;
            }

            if (!mergedCellRows.Contains(rowIndex))
            {
                mergedCellRows.Add(rowIndex);
            }
        }

        private static void AddOverwriteWarning(List<string> overwriteWarnings, int rowIndex, int colIndex, ImageCellAvailability availability)
        {
            if (overwriteWarnings == null || rowIndex < 1 || colIndex < 1)
            {
                return;
            }

            string contentType = availability == ImageCellAvailability.OverwriteImage
                ? "已有图片"
                : "已有文本";
            string warning = string.Format("第{0}行第{1}列{2}，已覆盖插入新图片", rowIndex, colIndex, contentType);

            if (!overwriteWarnings.Contains(warning))
            {
                overwriteWarnings.Add(warning);
            }
        }

        #endregion



        private void CloseProgressForm()
        {
            try
            {
                _progressReporter?.Close();
            }
            catch (Exception ex)
            {
                SafeIgnore(ex, "关闭进度窗口失败");
            }
        }

        private static void SafeIgnore(Exception ex, string context)
        {
            Debug.WriteLine($"{context}: {ex.Message}");
        }

    }
}


