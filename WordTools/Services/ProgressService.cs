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
        private int _statusBarUpdateInterval = 1;  // 状态栏更新间隔（每N张图片更新一次）
        private DateTime _lastStatusBarUpdate = DateTime.MinValue;  // 上次状态栏更新时间
        private const int STATUS_BAR_VISIBLE_MS = 500;  // 状态栏最小可见时间（毫秒），确保用户能看清

        // 性能模式备份
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;

        // 进度窗口
        private ProgressForm _progressForm;

        public ProgressService(Application application)
        {
            _application = application;
        }

        #region 取消控制

        /// <summary>
        /// 检查是否按下ESC键
        /// </summary>
        private bool CheckEscapeKey()
        {
            return (GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0;
        }

        /// <summary>
        /// 检查是否需要取消（支持ESC键和进度窗口取消按钮）
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
                _application.StatusBar = "检测到ESC键，正在取消操作...";
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
            if (totalFiles <= 10) return 1;     // 10张以内：每张都更新
            if (totalFiles <= 50) return 5;     // 50张以内：每5张更新
            if (totalFiles <= 200) return 10;   // 200张以内：每10张更新
            return 20;                           // 更多：每20张更新（最大化性能）
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
                    // 常规第0代快速回收
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
            _isCancelled = false;
            int processedCount = 0;
            int successCount = 0;
            int failCount = 0;
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
                // 优化：只在 startRow 位于现有内容范围内且 startRow 之后有实际内容时才清理
                // 如果 startRow 已经超出当前表格行数，说明是在末尾追加，新行为空无需清理
                int currentRowCount = tbl.Rows.Count;
                bool needClearNumbering = false;
                if (startRow <= currentRowCount)
                {
                    // 超轻量检查：只检查 startRow 行第1列的文本内容（不含InlineShapes.Count）
                    try
                    {
                        Range checkRange = tbl.Cell(startRow, 1).Range;
                        checkRange.SetRange(checkRange.Start, checkRange.End - 1);
                        string text = (checkRange.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        // 检查文本是否以数字开头（已有编号）或包含图片标记
                        bool hasContent = !string.IsNullOrEmpty(text);
                        bool hasPictureMarker = text.IndexOf('\x01') >= 0; // 图片在Range.Text中表示为\x01
                        
                        // 如果 startRow 有内容（文本或图片标记），才需要清理编号
                        if (hasContent || hasPictureMarker)
                        {
                            needClearNumbering = true;
                        }
                        // 注意：不再遍历 tbl.Range.InlineShapes 检查后续行，该操作在大表格中极慢（5-6秒）
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

                // 一次遍历获取图片文件列表和总数（避免重复扫描目录）
                var imageFiles = FileService.GetImageFiles(folderPath, includeRootImages, includeSubFolderImages);
                int totalFiles = imageFiles.TotalCount;

                if (totalFiles == 0)
                {
                    MessageBox.Show("未找到任何图片文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 显示提示（仅当图片数量超过20时）
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

                // 进入高性能模式（关闭 ScreenUpdating 以提升性能）
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
                int currentNumber = startNumber;

                // 预分配行数
                ImageService.PreAllocateRows(tbl, totalFiles, 2, needDescription, _application);
                t3 = sw.ElapsedMilliseconds;

                // 进度窗口已显示，无需状态栏提示

                // 获取文件夹名称
                string rootFolderName = FileService.GetFolderName(folderPath);

                // 处理根目录图片
                if (includeRootImages && imageFiles.RootFiles != null && imageFiles.RootFiles.Length > 0)
                {
                    // 创建标题行
                    TableService.CreateTitleRow(tbl, ref rowIndex, rootFolderName);

                    // 处理文件
                    ProcessFileBatch(imageFiles.RootFiles, tbl, ref rowIndex, minHeight, needDescription,
                        useFileNameAsDescription, useFolderNameAsDescription, rootFolderName, ref processedCount, ref successCount, ref failCount,
                        totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber);
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

                            ProcessFileBatch(subFiles, tbl, ref rowIndex, minHeight, needDescription,
                                useFileNameAsDescription, useFolderNameAsDescription, subfolderName, ref processedCount, ref successCount, ref failCount,
                                totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber);

                            // 释放引用，帮助垃圾回收
                            imageFiles.SubfolderFiles[subfolder] = null;
                        }
                    }
                }

                // 记录t4（图片插入完成时间）
                t4 = sw.ElapsedMilliseconds;

                // 显示完成消息
                stopwatch.Stop();
                double seconds = stopwatch.Elapsed.TotalSeconds;
                string timeInfo = seconds >= 60
                    ? $"{(int)(seconds / 60)}分{seconds % 60:F1}秒"
                    : $"{seconds:F1}秒";

                // 记录t5（最终时间）用于计算收尾耗时
                t5 = sw.ElapsedMilliseconds;

                string timeDetail = $"\n\n[诊断]\n" +
                    $"初始化: {t0}ms\n" +
                    $"清理编号: {t1 - t0}ms (跳过={skippedClear})\n" +
                    $"计算起始号: {t2 - t1}ms\n" +
                    $"预分配行: {t3 - t2}ms\n" +
                    $"插入图片: {t4 - t3}ms\n" +
                    $"收尾: {t5 - t4}ms";

                // 更新进度窗口为完成状态
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    _progressForm.ShowCompletion(successCount, failCount, seconds);
                }

                if (_isCancelled)
                {
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo) + timeDetail, "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (failCount > 0)
                {
                    MessageBox.Show(string.Format("图片插入完成！\n成功: {0} 张\n失败: {1} 张\n耗时: {2}", successCount, failCount, timeInfo) + timeDetail, "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo) + timeDetail, "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("处理过程中发生错误: {0}", ex.Message), "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

                // 1. 先退出高性能模式，让用户立即看到已插入的图片
                ExitHighPerformanceMode();

                // 循环DoEvents让Word有充足时间完成屏幕重绘
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
                    _progressForm.Close();
                    _progressForm = null;
                }

                _application.StatusBar = "";
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
            _isCancelled = false;
            int processedCount = 0;
            int successCount = 0;
            int failCount = 0;
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
                if (!TableService.IsTableFixedColumnWidth(tbl))
                {
                    TableService.SetTableFixedColumnWidth(tbl);
                }

                TableService.AdjustTableColumns(tbl, 2);
                int rowIndex = selection.Cells[1].RowIndex;
                startRow = rowIndex;

                int totalFiles = files.Length;
                _refreshInterval = GetOptimizedRefreshInterval(totalFiles);
                _statusBarUpdateInterval = GetStatusBarUpdateInterval(totalFiles);

                // 清除 startRow 之后的编号（增量模式，不影响前面已有的编号）
                // 优化：只在 startRow 位于现有内容范围内且 startRow 之后有实际内容时才清理
                // 如果 startRow 已经超出当前表格行数，说明是在末尾追加，新行为空无需清理
                int currentRowCount = tbl.Rows.Count;
                bool needClearNumbering = false;
                if (startRow <= currentRowCount)
                {
                    // 超轻量检查：只检查 startRow 行第1列的文本内容（不含InlineShapes.Count）
                    try
                    {
                        Range checkRange = tbl.Cell(startRow, 1).Range;
                        checkRange.SetRange(checkRange.Start, checkRange.End - 1);
                        string text = (checkRange.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        // 检查文本是否以数字开头（已有编号）或包含图片标记
                        bool hasContent = !string.IsNullOrEmpty(text);
                        bool hasPictureMarker = text.IndexOf('\x01') >= 0; // 图片在Range.Text中表示为\x01
                        
                        // 如果 startRow 有内容（文本或图片标记），才需要清理编号
                        if (hasContent || hasPictureMarker)
                        {
                            needClearNumbering = true;
                        }
                        // 注意：不再遍历 tbl.Range.InlineShapes 检查后续行，该操作在大表格中极慢（5-6秒）
                        // 如果用户从中间插入，需要清理编号，可通过其他方式检测（如检查描述行是否有编号文本）
                    }
                    catch { needClearNumbering = true; } // 出错时保守处理，执行清理
                }
                
                if (needClearNumbering)
                {
                    TableService.ClearTableNumbering(tbl, startRow);
                }

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
                int currentNumber = startNumber;

                // 处理文件
                string folderName = FileService.GetFolderName(FileService.GetParentFolder(files[0]));
                ProcessFileBatch(files, tbl, ref rowIndex, minHeight, needDescription,
                    useFileNameAsDescription, useFolderNameAsDescription, folderName, ref processedCount, ref successCount, ref failCount,
                    totalFiles, startTime, needAutoNumbering, wdAlignment, numberPosition, ref isFirstSeqField, ref currentNumber);

                stopwatch.Stop();
                double seconds = stopwatch.Elapsed.TotalSeconds;
                string timeInfo = seconds >= 60
                    ? $"{(int)(seconds / 60)}分{seconds % 60:F1}秒"
                    : $"{seconds:F1}秒";

                // 更新进度窗口为完成状态
                if (_progressForm != null && !_progressForm.IsDisposed)
                {
                    _progressForm.ShowCompletion(successCount, failCount, seconds);
                }

                if (_isCancelled)
                {
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。\n耗时: {1}", successCount, timeInfo), "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！\n耗时: {1}", successCount, timeInfo), "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("处理过程中发生错误: {0}", ex.Message), "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 1. 先退出高性能模式，让用户立即看到已插入的图片
                ExitHighPerformanceMode();

                // 循环DoEvents让Word有充足时间完成屏幕重绘
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
                    _progressForm.Close();
                    _progressForm = null;
                }

                _application.StatusBar = "";
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
            ref bool isFirstSeqField, ref int currentNumber)
        {
            var currentRowFiles = new List<string>();

            // 性能优化：缓存行数，避免重复 COM 调用
            int cachedRowCount = tbl.Rows.Count;

            // 性能优化：循环前预先调整列数
            TableService.AdjustTableColumns(tbl, 2);

            for (int i = 0; i < files.Length; i++)
            {
                if (ShouldCancel()) break;

                string filePath = files[i];
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
                    // 计算列位置
                    int expectedCol = (i % 2) + 1;

                    // 确保行存在（使用缓存行数优化）
                    if (rowIndex > cachedRowCount)
                    {
                        int rowCountBefore = cachedRowCount;
                        tbl.Rows.Add();
                        cachedRowCount = tbl.Rows.Count;
                        if (cachedRowCount <= rowCountBefore)
                        {
                            throw new InvalidOperationException("Failed to add new row to table.");
                        }
                    }

                    int colIndex = expectedCol;

                    // 检查单元格是否适合
                    if (!TableService.IsCellSuitableForImage(tbl.Cell(rowIndex, colIndex)))
                    {
                        int foundRow, foundCol;
                        if (TableService.FindNextSuitableCell(tbl, rowIndex, out foundRow, out foundCol, expectedCol))
                        {
                            rowIndex = foundRow;
                            colIndex = foundCol;
                        }
                        else
                        {
                            rowIndex = tbl.Rows.Count + 1;
                            tbl.Rows.Add();
                            colIndex = expectedCol;
                        }
                    }

                    // 插入图片
                    ImageService.InsertImageFast(tbl.Cell(rowIndex, colIndex), filePath, minHeight);
                    successCount++;

                    // 记录文件名或文件夹名
                    if (useFileNameAsDescription)
                    {
                        currentRowFiles.Add(filePath);
                    }
                    else if (useFolderNameAsDescription)
                    {
                        // 去掉文件夹名前面的序号
                        // 支持格式："1-名称"、"1.名称"、"01 名称"、"01. 名称"、"1 - 名称" 等
                        string cleanedFolderName = System.Text.RegularExpressions.Regex.Replace(folderName, @"^\d+\s*[\.\-]?\s*", "");
                        currentRowFiles.Add(cleanedFolderName);
                    }

                    // 行满处理
                    if ((i + 1) % 2 == 0 && i < files.Length - 1)
                    {
                        if (useFileNameAsDescription || useFolderNameAsDescription)
                        {
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                            // 保存当前行号，因为 InsertFileNameDescriptionRow 会递增 rowIndex
                            int descriptionRow = rowIndex;
                            bool descriptionsAreFilePaths = useFileNameAsDescription && !useFolderNameAsDescription;
                            TableService.InsertFileNameDescriptionRow(tbl, ref rowIndex, currentRowFiles.ToArray(), descriptionsAreFilePaths);
                            // 内联编号：直接插入纯文本编号，每列独立递增
                            if (needAutoNumbering)
                            {
                                for (int col = 1; col <= 2; col++)
                                {
                                    try { TableService.InsertNumberText(tbl, descriptionRow, col, numberAlignment, currentNumber, true, numberPosition); }
                                    catch { }
                                    currentNumber++; // 每列递增
                                }
                            }
                            currentRowFiles.Clear();
                        }
                        else if (needDescription)
                        {
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
                            TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                        }
                        else
                        {
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex, ref cachedRowCount);
                        }
                    }
                }
                catch
                {
                    failCount++;
                }

                processedCount++;

                // 定期清理内存
                if (processedCount % _memoryCleanInterval == 0)
                {
                    CleanupMemory(processedCount);
                }
            }

            // 处理最后一行（如果文件数量为奇数，需要填充空单元格）
            if (files.Length % 2 != 0)
            {
                TableService.FillEmptyCellsWithNA(tbl, rowIndex, 2, 2);
            }

            // 处理末尾的描述行
            if ((useFileNameAsDescription || useFolderNameAsDescription) && currentRowFiles.Count > 0)
            {
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
            }
            else if (needDescription)
            {
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

                // 确保 Word 窗口保持在前台可见
                EnsureWordWindowActive();

                // 进度由独立窗口显示，不再更新 Word 状态栏
            }
            catch
            {
                // 忽略错误
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
    }
}
