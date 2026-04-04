using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordTools.Services
{
    /// <summary>
    /// 进度服务
    /// 管理批量插图进度，提供性能优化支持
    /// </summary>
    public class ProgressService
    {
        // Windows API 声明（用于检测按键）
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);
        private const int VK_ESCAPE = 0x1B;

        private readonly Application _application;
        private bool _isCancelled;

        // 批量处理配置
        private int _refreshInterval = 10;
        private int _memoryCleanInterval = 50;
        private int _saveInterval = 200;

        // 性能模式备份
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;

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
        /// 检查是否需要取消
        /// </summary>
        private bool ShouldCancel()
        {
            if (_isCancelled) return true;

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
        /// 进入高性能模式
        /// </summary>
        private void EnterHighPerformanceMode()
        {
            try
            {
                _originalScreenUpdating = _application.ScreenUpdating;
                _originalDisplayAlerts = _application.DisplayAlerts != WdAlertLevel.wdAlertsNone;

                _application.ScreenUpdating = false;
                _application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                var doc = _application.ActiveDocument;
                if (doc != null)
                {
                    doc.SpellingChecked = true;
                    doc.GrammarChecked = true;
                }
            }
            catch
            {
                // 忽略错误
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
            if (totalFiles < 50) return 5;
            if (totalFiles < 200) return 10;
            if (totalFiles < 500) return 20;
            if (totalFiles < 1000) return 50;
            if (totalFiles < 2000) return 100;
            return 200;
        }

        /// <summary>
        /// 清理内存
        /// </summary>
        private void CleanupMemory()
        {
            try
            {
                System.Windows.Forms.Application.DoEvents();
                GC.Collect();
                GC.WaitForPendingFinalizers();
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
            bool needDescription, bool useFileNameAsDescription,
            bool includeRootImages, bool includeSubFolderImages,
            bool needAutoNumbering, int numberAlignment = 2)
        {
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

                // 清除编号
                TableService.ClearTableNumbering(tbl, startRow);

                // 计算总文件数
                int totalFiles = FileService.CountTotalImageFiles(folderPath, includeRootImages, includeSubFolderImages);

                if (totalFiles == 0)
                {
                    MessageBox.Show("未找到任何图片文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 显示提示
                MessageBox.Show(string.Format("开始插入图片...\n\n提示：插入过程中可以按 ESC 键随时取消操作。\n\n共找到 {0} 张图片。", totalFiles),
                    "批量插图", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 设置优化参数
                _refreshInterval = GetOptimizedRefreshInterval(totalFiles);
                _memoryCleanInterval = _refreshInterval * 5;
                _saveInterval = _refreshInterval * 20;

                // 进入高性能模式
                EnterHighPerformanceMode();

                // 预分配行数
                ImageService.PreAllocateRows(tbl, totalFiles, 2, needDescription, _application);

                _application.StatusBar = string.Format("准备插入 {0} 张图片...", totalFiles);
                System.Windows.Forms.Application.DoEvents();

                // 获取文件夹名称
                string rootFolderName = FileService.GetFolderName(folderPath);

                // 处理根目录图片
                if (includeRootImages)
                {
                    var rootFiles = FileService.GetRootImageFiles(folderPath);
                    if (rootFiles.Length > 0)
                    {
                        // 创建标题行
                        TableService.CreateTitleRow(tbl, ref rowIndex, rootFolderName);

                        // 处理文件
                        ProcessFileBatch(rootFiles, tbl, ref rowIndex, minHeight, needDescription,
                            useFileNameAsDescription, ref processedCount, ref successCount, ref failCount,
                            totalFiles, startTime);
                    }
                }

                // 处理子文件夹
                if (includeSubFolderImages)
                {
                    var subfolders = FileService.GetSubfolders(folderPath);
                    foreach (var subfolder in subfolders)
                    {
                        if (ShouldCancel()) break;

                        var subFiles = FileService.GetRootImageFiles(subfolder);
                        if (subFiles.Length > 0)
                        {
                            string subfolderName = FileService.GetFolderName(subfolder);
                            TableService.CreateTitleRow(tbl, ref rowIndex, subfolderName);

                            ProcessFileBatch(subFiles, tbl, ref rowIndex, minHeight, needDescription,
                                useFileNameAsDescription, ref processedCount, ref successCount, ref failCount,
                                totalFiles, startTime);
                        }
                    }
                }

                // 完成
                _application.StatusBar = string.Format("完成！成功: {0} 失败: {1}", successCount, failCount);

                // 显示完成消息
                if (_isCancelled)
                {
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。", successCount), "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (failCount > 0)
                {
                    MessageBox.Show(string.Format("图片插入完成！\n成功: {0} 张\n失败: {1} 张", successCount, failCount), "完成",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！", successCount), "完成",
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
                // 添加编号
                if (needAutoNumbering && startRow > 0 && tbl != null)
                {
                    try
                    {
                        TableService.AddNumberingToDescriptionRows(tbl, _application.ActiveDocument,
                            startRow, numberAlignment, needAutoNumbering);
                    }
                    catch
                    {
                        // 忽略编号错误
                    }
                }

                // 退出高性能模式
                ExitHighPerformanceMode();
                _application.StatusBar = "";
            }
        }

        #endregion

        #region 批量插入 - 选中文件

        /// <summary>
        /// 插入选中的图片（带进度）
        /// </summary>
        public void InsertSelectedPhotosWithProgress(string[] files, float minHeight,
            bool needDescription, bool useFileNameAsDescription, bool needAutoNumbering, int numberAlignment = 2)
        {
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

                EnterHighPerformanceMode();

                _application.StatusBar = string.Format("准备插入 {0} 张图片...", totalFiles);
                System.Windows.Forms.Application.DoEvents();

                // 处理文件
                ProcessFileBatch(files, tbl, ref rowIndex, minHeight, needDescription,
                    useFileNameAsDescription, ref processedCount, ref successCount, ref failCount,
                    totalFiles, startTime);

                // 完成
                _application.StatusBar = string.Format("完成！成功: {0} 失败: {1}", successCount, failCount);

                if (_isCancelled)
                {
                    MessageBox.Show(string.Format("操作已取消。已插入 {0} 张图片。", successCount), "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(string.Format("成功插入 {0} 张图片！", successCount), "完成",
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
                // 添加编号
                if (needAutoNumbering && startRow > 0 && tbl != null)
                {
                    try
                    {
                        TableService.AddNumberingToDescriptionRows(tbl, _application.ActiveDocument,
                            startRow, numberAlignment, needAutoNumbering);
                    }
                    catch
                    {
                        // 忽略编号错误
                    }
                }

                ExitHighPerformanceMode();
                _application.StatusBar = "";
            }
        }

        #endregion

        #region 文件批量处理

        /// <summary>
        /// 处理一批文件
        /// </summary>
        private void ProcessFileBatch(string[] files, Table tbl, ref int rowIndex,
            float minHeight, bool needDescription, bool useFileNameAsDescription,
            ref int processedCount, ref int successCount, ref int failCount,
            int totalFiles, DateTime startTime)
        {
            var currentRowFiles = new List<string>();

            for (int i = 0; i < files.Length; i++)
            {
                if (ShouldCancel()) break;

                string filePath = files[i];
                string fileName = FileService.GetFileName(filePath);

                // 更新进度
                if (processedCount % _refreshInterval == 0)
                {
                    UpdateStatusBar(processedCount, totalFiles, fileName, startTime);
                    System.Windows.Forms.Application.DoEvents();
                }

                try
                {
                    // 计算列位置
                    int expectedCol = (i % 2) + 1;

                    // 确保行存在
                    if (rowIndex > tbl.Rows.Count)
                    {
                        tbl.Rows.Add();
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

                    // 记录文件名
                    if (useFileNameAsDescription)
                    {
                        currentRowFiles.Add(filePath);
                    }

                    // 行满处理
                    if ((i + 1) % 2 == 0 && i < files.Length - 1)
                    {
                        if (useFileNameAsDescription)
                        {
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex);
                            TableService.InsertFileNameDescriptionRow(tbl, ref rowIndex, currentRowFiles.ToArray());
                            currentRowFiles.Clear();
                        }
                        else if (needDescription)
                        {
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex);
                            TableService.InsertDescriptionRow(tbl, ref rowIndex);
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex);
                        }
                        else
                        {
                            rowIndex++;
                            TableService.EnsureRowExists(tbl, rowIndex);
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
                    CleanupMemory();
                }
            }

            // 处理最后一行
            int batchSize = files.Length;
            if (batchSize % 2 != 0)
            {
                TableService.FillEmptyCellsWithNA(tbl, rowIndex, 2, 2);
            }

            // 处理末尾的描述行
            if (useFileNameAsDescription && currentRowFiles.Count > 0)
            {
                rowIndex++;
                TableService.EnsureRowExists(tbl, rowIndex);
                TableService.InsertFileNameDescriptionRow(tbl, ref rowIndex, currentRowFiles.ToArray());
            }
            else if (needDescription)
            {
                rowIndex++;
                TableService.EnsureRowExists(tbl, rowIndex);
                TableService.InsertDescriptionRow(tbl, ref rowIndex);
                rowIndex++;
            }
        }

        #endregion

        #region 状态栏更新

        /// <summary>
        /// 更新状态栏
        /// </summary>
        private void UpdateStatusBar(int current, int total, string currentFile, DateTime startTime)
        {
            try
            {
                int percent = total > 0 ? (int)((double)current / total * 100) : 0;
                var elapsed = DateTime.Now - startTime;
                double remaining = 0;

                if (current > 0)
                {
                    remaining = elapsed.TotalSeconds / current * (total - current);
                }

                string shortFileName = currentFile.Length > 30 
                    ? "..." + currentFile.Substring(currentFile.Length - 27) 
                    : currentFile;

                _application.StatusBar = string.Format("插入图片 {0}/{1} ({2}%) - 已用:{3:F0}s 剩余:{4:F0}s - {5}",
                    current, total, percent, elapsed.TotalSeconds, remaining, shortFileName);
            }
            catch
            {
                // 忽略错误
            }
        }

        #endregion
    }
}
