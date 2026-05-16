using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordTools.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordTools.Services
{
    /// <summary>
    /// 进度服务
    /// 管理批量插图进度，提供性能优化支持
    /// </summary>
    public class ProgressService
    {
        // Windows API 声明（用于检测按键和窗口置顶）
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);
        private const int VK_ESCAPE = 0x1B;

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;

        private readonly Application _application;
        private bool _isCancelled;

        // 批量处理配置
        private int _refreshInterval = 10;
        private int _memoryCleanInterval = 50;
        private int _saveInterval = 200;
        private int _fullGcInterval;  // 全代 GC 间隔
        private int _statusBarUpdateInterval = 1;  // 状态栏更新间隔（每 N 张图片更新一次）
        private DateTime _lastStatusBarUpdate = DateTime.MinValue;  // 上次状态栏更新时间
        private const int STATUS_BAR_VISIBLE_MS = 500;  // 状态栏最小可见时间（毫秒），确保用户能看清

        // 性能模式备份
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;

        // 进度窗口
        private ProgressForm _progressForm;
        private InsertionPerformanceDiagnostics _activeDiagnostics;

        public ProgressService(Application application)
        {
            _application = application;
        }

        #region 取消控制

        /// <summary>
        /// 检查是否按下 ESC 键
        /// </summary>
        private bool CheckEscapeKey()
        {
            return (GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0;
        }

        /// <summary>
        /// 检查是否需要取消（支持 ESC 键和进度窗口取消按钮）
        /// </summary>
        private bool ShouldCancel()
        {
            if (_isCancelled) return true;

            // 检查进度窗口是否点击了取消
            if (_progressForm != null && !_progressForm.IsDisposed && _progressForm.IsCancelled)
            {
                _isCancelled = true;
                return true;
            }

            if (CheckEscapeKey())
            {
                _isCancelled = true;
                _application.StatusBar = "检测到 ESC 键，正在取消操作...";
                System.Windows.Forms.Application.DoEvents();
                return true;
            }

            return false;
        }

        #endregion

        #region 性能优化

        /// <summary>
        /// 进入高性能模式（关闭 ScreenUpdating 以最大化插图性能）
        /// </summary>
        private void EnterHighPerformanceMode()
        {
            try
            {
                _originalScreenUpdating = _application.ScreenUpdating;
                _originalDisplayAlerts = _application.DisplayAlerts != WdAlertLevel.wdAlertsNone;

                // 关闭 ScreenUpdating 以提升插图性能（进度由独立窗口显示）
                _application.ScreenUpdating = false;
                _application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                var doc = _application.ActiveDocument;
                if (doc != null)
                {
                    doc.SpellingChecked = true;
                    doc.GrammarChecked = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProgressService] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// 退出高性能模式
        /// </summary>
        private void ExitHighPerformanceMode()
        {
            try
            {
                _application.ScreenUpdating = _originalScreenUpdating;
                _application.DisplayAlerts = _originalDisplayAlerts 
                    ? WdAlertLevel.wdAlertsAll 
                    : WdAlertLevel.wdAlertsNone;
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 根据文件数量获取优化的刷新间隔
        /// </summary>
        private int GetOptimizedRefreshInterval(int totalFiles)
        {
            if (totalFiles < 30) return 10;
            if (totalFiles < 100) return 15;
            return 20;
        }

        /// <summary>
        /// 根据文件数量获取进度更新间隔
        /// 大量图片时降低更新频率以最大化插图性能
        /// </summary>
        private int GetStatusBarUpdateInterval(int totalFiles)
        {
            if (totalFiles <= 10) return 1;     // 10 张以内：每张都更新
            if (totalFiles <= 50) return 5;     // 50 张以内：每 5 张更新
            if (totalFiles <= 200) return 15;   // 200 张以内：每 15 张更新
            return 25;                           // 更多：每 25 张更新（最大化性能）
        }

        /// <summary>
        /// 清理内存（分级 GC 策略）
        /// </summary>
        private void CleanupMemory(int processedCount)
        {
            try
            {
                System.Windows.Forms.Application.DoEvents();

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
            catch
            {
                // 忽略错误
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
            _isCancelled = false;
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
                var selection = _application.Selection;
                if (!TableService.IsSelectionInTable(selection))
                {
                    MessageBox.Show("请先选中一个表格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!TableService.IsSelectionInFirstColumn(selection))
                {
                    MessageBox.Show("请将光标置于表格左侧单元格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    catch { needClearNumbering = true; } // 出错时保守处理，执行清理
                }
                
                if (needClearNumbering)
                {
                    TableService.ClearTableNumbering(tbl, startRow);
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
                    MessageBox.Show("未找到任何图片文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 仅在图片数量较多时显示开始提示
                if (totalFiles > 20)
                {
                    MessageBox.Show(string.Format("开始插入图片...\n\n提示：插入过程中可以按 ESC 键随时取消操作。\n\n共找到 {0} 张图片。", totalFiles),
                        "批量插图", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // 设置优化参数
                _refreshInterval = GetOptimizedRefreshInterval(totalFiles);
                _statusBarUpdateInterval = GetStatusBarUpdateInterval(totalFiles);
                _memoryCleanInterval = _refreshInterval * 10;
                _fullGcInterval = _memoryCleanInterval * 5;
                _saveInterval = _refreshInterval * 20;

                // 进入高性能模式
                EnterHighPerformanceMode();

                // 创建并显示进度窗口
                _progressForm = new ProgressForm(totalFiles);
                _progressForm.Show();
                System.Windows.Forms.Application.DoEvents();

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
                    startNumber = TableService.CalculateNextSequenceNumber(tbl, startRow);
                }
                t2 = sw.ElapsedMilliseconds;
                int currentNumber = startNumber;

                // 预分配行数
                ImageService.PreAllocateRows(tbl, totalFiles, 2, needDescription, _application);
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
                        if (ShouldCancel()) break;

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

                string timeDetail = BuildTimeDetail(
                    showDetailedLog,
                    t0,
                    t1,
                    t2,
                    t3,
                    t4,
                    t5,
                    skippedClear,
                    insertionDiagnostics);

                // 更新进度窗口为完成状态
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    CloseProgressForm();
                }

                if (_isCancelled)
                {
                    benchmarkStatus = "Cancelled";
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo) + timeDetail, "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (InsertionSummaryFormatter.ShouldShowSummary(failCount, mergedCellRows, overwriteWarnings))
                {
                    benchmarkStatus = failCount > 0 ? "CompletedWithFailures" : "CompletedWithWarnings";
                    ShowInsertionSummary(successCount, failCount, timeInfo, timeDetail, failedFiles, mergedCellRows, overwriteWarnings);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo) + timeDetail, "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                benchmarkStatus = "Error";
                benchmarkError = ex.Message;
                MessageBox.Show(string.Format("处理过程中发生错误: {0}", ex.Message), "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (_isCancelled)
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
                ExitHighPerformanceMode();

                // 循环 DoEvents 让 Word 有充足时间完成屏幕重绘
                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(10);
                    }
                }
                catch { }

                // 3. 确保域代码不可见（保持原有逻辑不变）
                try
                {
                    if (_application.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
                    {
                        _application.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch { }

                // 关闭进度窗口
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    CloseProgressForm();
                }

                _application.StatusBar = "";

                TryWriteBenchmarkLog(new BenchmarkLogEntry
                {
                    RunMode = "Folder",
                    Status = benchmarkStatus,
                    SourcePath = folderPath,
                    TotalFiles = totalFiles,
                    ProcessedCount = processedCount,
                    SuccessCount = successCount,
                    FailCount = failCount,
                    MergedCellCount = mergedCellRows.Count,
                    Cancelled = _isCancelled,
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
            _isCancelled = false;
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
                var selection = _application.Selection;
                if (!TableService.IsSelectionInTable(selection))
                {
                    MessageBox.Show("请先选中一个表格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!TableService.IsSelectionInFirstColumn(selection))
                {
                    MessageBox.Show("请将光标置于表格左侧单元格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                _refreshInterval = GetOptimizedRefreshInterval(totalFiles);
                _statusBarUpdateInterval = GetStatusBarUpdateInterval(totalFiles);
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
                    catch { needClearNumbering = true; } // 出错时保守处理，执行清理
                }
                
                if (needClearNumbering)
                {
                    TableService.ClearTableNumbering(tbl, startRow);
                }
                else
                {
                    skippedClear = true;
                }
                t1 = sw.ElapsedMilliseconds;

                EnterHighPerformanceMode();

                // 创建并显示进度窗口（独立于 Word，不受 ScreenUpdating 影响）
                _progressForm = new ProgressForm(totalFiles);
                _progressForm.Show();
                System.Windows.Forms.Application.DoEvents();

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
                    startNumber = TableService.CalculateNextSequenceNumber(tbl, startRow);
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
                string timeDetail = BuildTimeDetail(
                    showDetailedLog,
                    t0,
                    t1,
                    t2,
                    t3,
                    t4,
                    t5,
                    skippedClear,
                    insertionDiagnostics);

                // 更新进度窗口为完成状态
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    CloseProgressForm();
                }

                if (_isCancelled)
                {
                    benchmarkStatus = "Cancelled";
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo) + timeDetail, "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (InsertionSummaryFormatter.ShouldShowSummary(failCount, mergedCellRows, overwriteWarnings))
                {
                    benchmarkStatus = failCount > 0 ? "CompletedWithFailures" : "CompletedWithWarnings";
                    ShowInsertionSummary(successCount, failCount, timeInfo, timeDetail, failedFiles, mergedCellRows, overwriteWarnings);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo) + timeDetail, "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                benchmarkStatus = "Error";
                benchmarkError = ex.Message;
                MessageBox.Show(string.Format("处理过程中发生错误: {0}", ex.Message), "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (_isCancelled)
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
                ExitHighPerformanceMode();

                // 循环 DoEvents 让 Word 有充足时间完成屏幕重绘
                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(10);
                    }
                }
                catch { }

                // 3. 确保域代码不可见（保持原有逻辑不变）
                try
                {
                    if (_application.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
                    {
                        _application.ActiveDocument.ActiveWindow.View.ShowFieldCodes = false;
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch { }

                // 关闭进度窗口
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    CloseProgressForm();
                }

                _application.StatusBar = "";

                TryWriteBenchmarkLog(new BenchmarkLogEntry
                {
                    RunMode = "SelectedFiles",
                    Status = benchmarkStatus,
                    SourcePath = sourcePath,
                    TotalFiles = files != null ? files.Length : 0,
                    ProcessedCount = processedCount,
                    SuccessCount = successCount,
                    FailCount = failCount,
                    MergedCellCount = mergedCellRows.Count,
                    Cancelled = _isCancelled,
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
                if (ShouldCancel()) break;

                string filePath = validFiles[i];
                string fileName = FileService.GetFileName(filePath);

                // 更新进度窗口
                if (processedCount % _statusBarUpdateInterval == 0)
                {
                    UpdateStatusBar(processedCount, totalFiles, fileName, startTime);
                    // 处理消息队列，确保进度窗口 UI 更新
                    System.Windows.Forms.Application.DoEvents();
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
                                        try { TableService.InsertNumberText(tbl, descriptionRow, col, numberAlignment, currentNumber, true, numberPosition); }
                                        catch { }
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
                                        try { TableService.InsertNumberText(tbl, rowIndex, col, numberAlignment, currentNumber, true, numberPosition); }
                                        catch { }
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

                        if (imagesPlacedInCurrentRow == 1 && IsMergedCellError(errorMsg))
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
                    string classifiedError = ClassifyInsertionError(ex);
                    failedFiles?.Add((fileName, classifiedError));

                    if (imagesPlacedInCurrentRow == 1 && IsMergedCellError(classifiedError))
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
                        try { TableService.InsertNumberText(tbl, descriptionRow, col, numberAlignment, currentNumber, true, numberPosition); }
                        catch { }
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
                        try { TableService.InsertNumberText(tbl, rowIndex, col, numberAlignment, currentNumber, true, numberPosition); }
                        catch { }
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
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    _progressForm.UpdateProgress(current, total, shortFileName, elapsed);
                    // 确保进度窗口保持置顶
                    EnsureWindowTopMost(_progressForm.Handle);
                }

                // 进度由独立窗口显示，不再更新 Word 状态栏
            }
            catch
            {
                // 忽略错误
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

        /// <summary>
        /// 确保窗口保持置顶（使用 Win32 API，比 TopMost 更可靠）
        /// </summary>
        private void EnsureWindowTopMost(IntPtr hWnd)
        {
            try
            {
                if (hWnd != IntPtr.Zero && IsWindow(hWnd))
                {
                    SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }
            catch { }
        }

        /// <summary>
        /// 确保 Word 窗口保持激活状态
        /// </summary>
        private void EnsureWordWindowActive()
        {
            try
            {
                if (_application != null && _application.ActiveWindow != null)
                {
                    _application.ActiveWindow.Activate();
                }
            }
            catch { }
        }

        #endregion

        #region 错误分类

        /// <summary>
        /// 将异常分类为用户友好的错误信息
        /// </summary>
        private string ClassifyInsertionError(Exception ex)
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
                return msg; // 直接使用预生成的描述
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

        private static bool IsMergedCellError(string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage))
            {
                return false;
            }

            return errorMessage.IndexOf("合并", StringComparison.OrdinalIgnoreCase) >= 0
                || errorMessage.IndexOf("merge", StringComparison.OrdinalIgnoreCase) >= 0;
        }

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

        #region 失败信息汇总

        /// <summary>
        /// 显示插图失败汇总信息，支持查看全部详情，同时显示合并单元格绕开信息
        /// </summary>
        private void ShowInsertionSummary(int successCount, int failCount, string timeInfo, string timeDetail,
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

            MessageBoxButtons buttons = showDetails ? MessageBoxButtons.YesNo : MessageBoxButtons.OK;
            var result = MessageBox.Show(
                summaryText,
                "插图完成",
                buttons,
                MessageBoxIcon.Information,
                showDetails ? MessageBoxDefaultButton.Button2 : MessageBoxDefaultButton.Button1);

            if (result == DialogResult.Yes && showDetails)
            {
                using (var detailsForm = new FailureDetailsForm(failedFiles, mergedCellRows, overwriteWarnings))
                {
                    detailsForm.ShowDialog();
                }
            }
        }

        private static string BuildTimeDetail(
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

        private void ShowFailureSummary(int successCount, int failCount, string timeInfo, string timeDetail,
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

            var result = MessageBox.Show(sb.ToString(), "插图完成",
                buttons,
                MessageBoxIcon.Information,
                buttons == MessageBoxButtons.YesNo ? MessageBoxDefaultButton.Button2 : MessageBoxDefaultButton.Button1);

            if (result == DialogResult.Yes && (hasMoreFailures || hasMoreMerged))
            {
                using (var detailsForm = new FailureDetailsForm(failedFiles, mergedCellRows))
                {
                    detailsForm.ShowDialog();
                }
            }
        }

        /// <summary>
        /// 显示合并单元格绕开提示（无失败时单独调用）
        /// </summary>
        private void ShowMergedCellWarning(List<int> mergedCellRows)
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
                var result = MessageBox.Show(sb.ToString(), "合并单元格提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    using (var detailsForm = new FailureDetailsForm(null, mergedCellRows))
                    {
                        detailsForm.ShowDialog();
                    }
                }
            }
            else
            {
                MessageBox.Show(sb.ToString(), "合并单元格提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ShowOverwriteWarning(List<string> overwriteWarnings)
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
                var result = MessageBox.Show(sb.ToString(), "覆盖插图提示",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    using (var detailsForm = new FailureDetailsForm(null, null, overwriteWarnings))
                    {
                        detailsForm.ShowDialog();
                    }
                }
            }
            else
            {
                MessageBox.Show(sb.ToString(), "覆盖插图提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void TryWriteBenchmarkLog(BenchmarkLogEntry entry)
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
                    documentPath = _application != null && _application.ActiveDocument != null
                        ? _application.ActiveDocument.FullName
                        : null;
                }
                catch
                {
                    documentPath = null;
                }

                string logPath = BenchmarkLogService.GetDefaultLogPath(documentPath);
                entry.DocumentPath = documentPath;
                entry.LogPath = logPath;
                BenchmarkLogService.AppendCsv(logPath, entry);
            }
            catch
            {
                // 基准日志仅用于开发调试，禁止影响主流程
            }
        }

        private void CloseProgressForm()
        {
            if (_progressForm == null)
            {
                return;
            }

            try
            {
                if (!_progressForm.IsDisposed)
                {
                    _progressForm.TopMost = false;
                    _progressForm.Close();
                }
            }
            catch
            {
            }
            finally
            {
                _progressForm = null;
            }
        }

        #endregion
    }
}


