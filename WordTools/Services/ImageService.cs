using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services
{
    /// <summary>
    /// 图片工具服务
    /// 处理图片插入和尺寸调整
    /// </summary>
    public static class ImageService
    {
        // 厘米到磅的转换系数
        private const float CM_TO_POINTS = 28.35f;

        // 最小可接受行高（磅），低于此值视为异常单行
        private const float MIN_ACCEPTABLE_ROW_HEIGHT = 10f;

        // COM 调用重试次数
        private const int COM_RETRY_COUNT = 3;

        // COM 调用重试间隔（毫秒）
        private const int COM_RETRY_DELAY_MS = 200;

        #region 尺寸转换

        /// <summary>
        /// 将厘米转换为磅
        /// </summary>
        /// <param name="heightCM">高度（厘米）</param>
        /// <returns>高度（磅）</returns>
        public static float ConvertCMToPoints(float heightCM)
        {
            return heightCM * CM_TO_POINTS;
        }

        /// <summary>
        /// 将磅转换为厘米
        /// </summary>
        /// <param name="heightPoints">高度（磅）</param>
        /// <returns>高度（厘米）</returns>
        public static float ConvertPointsToCM(float heightPoints)
        {
            return heightPoints / CM_TO_POINTS;
        }

        /// <summary>
        /// 验证并转换高度输入
        /// </summary>
        /// <param name="heightInput">用户输入的高度字符串（厘米）</param>
        /// <param name="heightPoints">转换后的高度（磅）</param>
        /// <returns>True 如果输入有效</returns>
        public static bool ValidateAndConvertHeight(string heightInput, out float heightPoints)
        {
            heightPoints = -1;

            if (string.IsNullOrEmpty(heightInput))
            {
                return true; // 空输入视为有效（不使用最小高度）
            }

            float tempHeight;
            if (!float.TryParse(heightInput, out tempHeight) || tempHeight <= 0)
            {
                return false;
            }

            heightPoints = ConvertCMToPoints(tempHeight);
            return true;
        }

        #endregion

        #region 单元格验证

        /// <summary>
        /// 验证单元格是否适合插入图片（增强版）
        /// 检查：合并单元格、行高异常、浮动图片、文件有效性
        /// </summary>
        public static bool ValidateCellForImage(Cell targetCell, string imagePath, out string errorMessage)
        {
            return ValidateCellForImage(targetCell, imagePath, out errorMessage, null);
        }

        public static bool ValidateCellForImage(Cell targetCell, string imagePath, out string errorMessage, ImageInsertionBatchContext context)
        {
            errorMessage = null;
            var validationWatch = context != null
                ? System.Diagnostics.Stopwatch.StartNew()
                : null;

            if (targetCell == null)
            {
                errorMessage = "目标单元格为空";
                return false;
            }

            try
            {
                // 1. 检查文件有效性
                if (!FileService.ValidateImageFile(imagePath, out errorMessage))
                {
                    return false;
                }

                // 2. 检查合并单元格（合并单元格的 Cell 对象行为异常）
                try
                {
                    // 尝试访问 RowIndex 和 ColumnIndex，合并单元格可能抛出异常或返回异常值
                    int rowIdx = targetCell.RowIndex;
                    int colIdx = targetCell.ColumnIndex;
                    if (rowIdx < 1 || colIdx < 1)
                    {
                        errorMessage = "单元格索引异常（可能是合并单元格）";
                        return false;
                    }
                }
                catch (COMException)
                {
                    errorMessage = "无法访问单元格（可能是合并单元格）";
                    return false;
                }

                // 3. 检查行高是否异常（单行压缩）
                try
                {
                    float rowHeight = targetCell.Height;
                    if (rowHeight > 0 && rowHeight < MIN_ACCEPTABLE_ROW_HEIGHT)
                    {
                        errorMessage = string.Format("表格行高异常（{0:F1}磅），请调整行高后再插入", rowHeight);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("行高读取失败，继续其他检查: " + ex.Message);
                }

                // 4. 检查单元格宽度是否异常
                try
                {
                    float cellWidth = targetCell.Width;
                    if (cellWidth > 0 && cellWidth < 5)
                    {
                        errorMessage = string.Format("单元格宽度异常（{0:F1}磅），请调整列宽后再插入", cellWidth);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("宽度读取失败，继续其他检查: " + ex.Message);
                }

                return true;
            }
            catch (COMException ex)
            {
                errorMessage = "COM 错误: " + ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "单元格验证失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (validationWatch != null)
                {
                    validationWatch.Stop();
                    context.Diagnostics.RecordCellValidation(validationWatch.ElapsedMilliseconds);
                }
            }
        }

        #endregion

        #region COM 重试

        /// <summary>
        /// 带重试机制的 AddPicture 调用
        /// 处理 Word COM 繁忙时的 RPC_E_SERVERCALL_RETRYLATER 错误
        /// </summary>
        private static InlineShape TryAddPictureWithRetry(Range targetRange, string imagePath, out string errorMessage)
        {
            errorMessage = null;
            InlineShape shape = null;

            for (int attempt = 1; attempt <= COM_RETRY_COUNT; attempt++)
            {
                try
                {
                    shape = targetRange.InlineShapes.AddPicture(
                        FileName: imagePath,
                        LinkToFile: false,
                        SaveWithDocument: true);

                    // 验证插入后的尺寸是否有效
                    if (shape.Width <= 0 || shape.Height <= 0)
                    {
                        errorMessage = "图片尺寸异常（Width=" + shape.Width + ", Height=" + shape.Height + "）";
                        try { shape.Delete(); } catch (Exception deleteEx) { System.Diagnostics.Debug.WriteLine("删除异常尺寸图片失败: " + deleteEx.Message); }
                        return null;
                    }

                    return shape;
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80010001) ||  // RPC_E_CALL_REJECTED
                                               ex.HResult == unchecked((int)0x8001010A) ||  // RPC_E_SERVERCALL_RETRYLATER
                                               ex.Message.IndexOf("rejected", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                               ex.Message.IndexOf("retry", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                               ex.Message.IndexOf("忙", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                               ex.Message.IndexOf("busy", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    errorMessage = "Word 正忙，无法插入图片";
                    if (attempt < COM_RETRY_COUNT)
                    {
                        System.Threading.Thread.Sleep(COM_RETRY_DELAY_MS * attempt);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch (COMException ex)
                {
                    errorMessage = "COM 错误: " + ex.Message;
                    break;
                }
                catch (FileNotFoundException)
                {
                    errorMessage = "图片文件未找到";
                    break;
                }
                catch (IOException ex)
                {
                    errorMessage = "文件访问错误: " + ex.Message;
                    break;
                }
                catch (Exception ex)
                {
                    errorMessage = "插入异常: " + ex.Message;
                    break;
                }
            }

            return null;
        }

        private static void ClearCellContentForOverwrite(Cell targetCell)
        {
            if (targetCell == null)
            {
                return;
            }

            try
            {
                Range cellRange = targetCell.Range;
                for (int i = cellRange.InlineShapes.Count; i >= 1; i--)
                {
                    try
                    {
                        cellRange.InlineShapes[i].Delete();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("删除单元格内嵌图片失败: " + ex.Message);
                    }
                }

                try
                {
                    var shapes = cellRange.Document.Shapes;
                    for (int i = shapes.Count; i >= 1; i--)
                    {
                        try
                        {
                            var shape = shapes[i];
                            if (shape != null && shape.Anchor != null &&
                                IsShapeAnchorWithinCell(shape.Anchor.Start, shape.Anchor.End, cellRange.Start, cellRange.End))
                            {
                                shape.Delete();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("删除单元格浮动图片失败: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("扫描单元格浮动图片失败: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("清空单元格内容失败: " + ex.Message);
            }

            targetCell.Range.Text = "";
        }

        private static void ClearCellContentForOverwrite(Cell targetCell, ImageInsertionBatchContext context)
        {
            if (context == null)
            {
                ClearCellContentForOverwrite(targetCell);
                return;
            }

            var clearWatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                if (targetCell == null)
                {
                    return;
                }

                Range cellRange = targetCell.Range;
                for (int i = cellRange.InlineShapes.Count; i >= 1; i--)
                {
                    try
                    {
                        cellRange.InlineShapes[i].Delete();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("删除单元格内嵌图片失败: " + ex.Message);
                    }
                }

                bool shouldScanFloatingShapes = true;
                try
                {
                    FloatingShapeIndex index = context.GetOrCreateFloatingShapeIndex(
                        () => FloatingShapeIndex.Create(CollectFloatingShapeAnchors(cellRange.Document.Shapes)));
                    shouldScanFloatingShapes = index.HasShapeInRange(cellRange.Start, cellRange.End);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("浮动图片索引构建失败，回退到全量扫描: " + ex.Message);
                    shouldScanFloatingShapes = true;
                }

                if (shouldScanFloatingShapes)
                {
                    try
                    {
                        var shapes = cellRange.Document.Shapes;
                        for (int i = shapes.Count; i >= 1; i--)
                        {
                            try
                            {
                                var shape = shapes[i];
                                if (shape != null && shape.Anchor != null &&
                                    IsShapeAnchorWithinCell(shape.Anchor.Start, shape.Anchor.End, cellRange.Start, cellRange.End))
                                {
                                    shape.Delete();
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine("删除单元格浮动图片失败: " + ex.Message);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("扫描单元格浮动图片失败: " + ex.Message);
                    }
                }

                targetCell.Range.Text = "";
            }
            finally
            {
                clearWatch.Stop();
                context.Diagnostics.RecordOverwriteClear(clearWatch.ElapsedMilliseconds);
                context.InvalidateFloatingShapeIndex();
            }
        }

        private static List<FloatingShapeAnchor> CollectFloatingShapeAnchors(Shapes shapes)
        {
            var anchors = new List<FloatingShapeAnchor>();
            if (shapes == null)
            {
                return anchors;
            }

            for (int i = 1; i <= shapes.Count; i++)
            {
                Shape shape = null;
                try
                {
                    shape = shapes[i];
                    if (shape != null && shape.Anchor != null)
                    {
                        anchors.Add(new FloatingShapeAnchor(shape.Anchor.Start, shape.Anchor.End));
                    }
                }
                catch
                {
                }
            }

            return anchors;
        }

        public static bool IsShapeAnchorWithinCell(int anchorStart, int anchorEnd, int cellStart, int cellEnd)
        {
            return anchorStart >= cellStart && anchorEnd <= cellEnd;
        }

        #endregion

        #region 图片插入

        /// <summary>
        /// 插入图片到单元格并调整尺寸
        /// </summary>
        /// <param name="targetCell">目标单元格</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="errorMessage">错误信息输出</param>
        /// <param name="minHeightPoints">最小高度（磅），-1 表示不限制</param>
        /// <returns>插入的图片对象</returns>
        public static InlineShape InsertImageToCell(Cell targetCell, string imagePath, out string errorMessage, float minHeightPoints = -1)
        {
            errorMessage = null;

            // 增强验证
            if (!ValidateCellForImage(targetCell, imagePath, out errorMessage))
            {
                return null;
            }

            try
            {
                // 获取单元格的宽度和高度
                float cellWidth = targetCell.Width;
                float cellHeight = targetCell.Height;

                // 计算目标尺寸（考虑边距）
                float targetWidth = cellWidth - 6; // 减去左右边距各3磅
                float targetHeight = cellHeight - 6; // 减去上下边距各3磅

                // 尺寸异常时使用安全值
                if (targetWidth < 1) targetWidth = 1;
                if (targetHeight < 1) targetHeight = 1;

                // 插入图片（带重试）
                var p = TryAddPictureWithRetry(targetCell.Range, imagePath, out errorMessage);
                if (p == null)
                {
                    return null;
                }

                // msoTrue = -1
                ((dynamic)p).LockAspectRatio = -1;

                // 计算缩放比例，保持宽高比
                float scaleRatio = 1;

                // 根据宽度限制计算缩放比例
                if (p.Width > targetWidth)
                {
                    scaleRatio = targetWidth / p.Width;
                }

                // 根据高度限制调整缩放比例
                if (p.Height * scaleRatio > targetHeight)
                {
                    scaleRatio = targetHeight / p.Height;
                }

                // 应用缩放（同时设置 Width 和 Height，不依赖 LockAspectRatio 的自动联动）
                if (scaleRatio < 1)
                {
                    p.Width = p.Width * scaleRatio;
                    p.Height = p.Height * scaleRatio;
                }

                // 如果设置了最小高度且当前高度小于最小高度，安全调整
                if (minHeightPoints > 0 && p.Height < minHeightPoints)
                {
                    float ratio = minHeightPoints / p.Height;
                    ((dynamic)p).LockAspectRatio = 0;
                    p.Height = minHeightPoints;
                    p.Width = p.Width * ratio;
                    ((dynamic)p).LockAspectRatio = -1;
                }

                return p;
            }
            catch (Exception ex)
            {
                errorMessage = "插图失败: " + ex.Message;
                System.Diagnostics.Debug.WriteLine($"InsertImageToCell error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 快速插入图片（最小化内存使用，优化性能）
        /// </summary>
        /// <param name="targetCell">目标单元格</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="errorMessage">错误信息输出</param>
        /// <param name="minHeightPoints">最小高度（磅），-1 表示不限制</param>
        /// <returns>是否插入成功</returns>
        public static bool InsertImageFast(Cell targetCell, string imagePath, out string errorMessage, float minHeightPoints = -1)
        {
            return InsertImageFast(targetCell, imagePath, out errorMessage, minHeightPoints, null);
        }

        public static bool InsertImageFast(Cell targetCell, string imagePath, out string errorMessage, float minHeightPoints, ImageInsertionBatchContext context)
        {
            errorMessage = null;

            // 增强验证
            if (!ValidateCellForImage(targetCell, imagePath, out errorMessage, context))
            {
                return false;
            }

            try
            {
                // 缓存单元格尺寸，避免重复 COM 调用
                float cellWidth = targetCell.Width - 6;
                float cellHeight = targetCell.Height - 6;

                // 尺寸异常时使用安全值，避免缩放比例计算错误
                if (cellWidth < 1) cellWidth = 1;
                if (cellHeight < 1) cellHeight = 1;

                // 清空单元格现有内容，覆盖旧文本/旧图片
                ClearCellContentForOverwrite(targetCell, context);

                // 插入图片（带重试）
                var addPictureWatch = System.Diagnostics.Stopwatch.StartNew();
                var p = TryAddPictureWithRetry(targetCell.Range, imagePath, out errorMessage);
                addPictureWatch.Stop();
                if (context != null)
                {
                    context.Diagnostics.RecordAddPicture(addPictureWatch.ElapsedMilliseconds);
                }
                if (p == null)
                {
                    return false;
                }

                var sizingWatch = context != null
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                try
                {
                    // 设置锁定宽高比（msoTrue = -1）
                    ((dynamic)p).LockAspectRatio = -1;

                    // 一次性计算并应用缩放（减少 COM 交互次数）
                    float scaleRatio = 1.0f;

                    // 根据宽度限制计算缩放比例
                    if (p.Width > cellWidth)
                    {
                        scaleRatio = cellWidth / p.Width;
                    }

                    // 根据高度限制调整缩放比例
                    float scaledHeight = p.Height * scaleRatio;
                    if (cellHeight > 10 && scaledHeight > cellHeight)
                    {
                        scaleRatio = cellHeight / p.Height;
                    }

                    // 应用缩放（同时设置 Width 和 Height，不依赖 LockAspectRatio 的自动联动）
                    if (scaleRatio < 1.0f)
                    {
                        p.Width = p.Width * scaleRatio;
                        p.Height = p.Height * scaleRatio;
                    }

                    // 最小高度限制（安全处理：临时解除锁定，按比例调整宽高）
                    if (minHeightPoints > 0 && p.Height < minHeightPoints)
                    {
                        float ratio = minHeightPoints / p.Height;
                        ((dynamic)p).LockAspectRatio = 0;
                        p.Height = minHeightPoints;
                        p.Width = p.Width * ratio;
                        ((dynamic)p).LockAspectRatio = -1;
                    }
                }
                finally
                {
                    if (sizingWatch != null)
                    {
                        sizingWatch.Stop();
                        context.Diagnostics.RecordPictureSizing(sizingWatch.ElapsedMilliseconds);
                    }
                }

                // 释放 COM 对象引用
                Marshal.ReleaseComObject(p);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "插图失败: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region 批量操作

        /// <summary>
        /// 批量调整已插入图片尺寸
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endRow">结束行</param>
        /// <param name="minHeightPoints">最小高度（磅）</param>
        public static void BatchResizeImages(Table tbl, int startRow, int endRow, float minHeightPoints = -1)
        {
            if (tbl == null) return;

            try
            {
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = 1; col <= tbl.Columns.Count; col++)
                    {
                        try
                        {
                            var cell = tbl.Cell(row, col);

                            // 检查单元格中是否有图片
                            if (cell.Range.InlineShapes.Count > 0)
                            {
                                foreach (InlineShape shape in cell.Range.InlineShapes)
                                {
                                    // msoTrue = -1
                                    ((dynamic)shape).LockAspectRatio = -1;

                                    // 调整尺寸
                                    float availableWidth = cell.Width - 6;
                                    if (availableWidth < 1) availableWidth = 1;

                                    if (shape.Width > availableWidth)
                                    {
                                        float ratio = availableWidth / shape.Width;
                                        shape.Width = availableWidth;
                                        shape.Height = shape.Height * ratio;
                                    }

                                    if (minHeightPoints > 0 && shape.Height < minHeightPoints)
                                    {
                                        float ratio = minHeightPoints / shape.Height;
                                        ((dynamic)shape).LockAspectRatio = 0;
                                        shape.Height = minHeightPoints;
                                        shape.Width = shape.Width * ratio;
                                        ((dynamic)shape).LockAspectRatio = -1;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // 忽略单个单元格错误
                        }
                    }
                }
            }
            catch
            {
                // 忽略批量操作错误
            }
        }

        /// <summary>
        /// 批量添加行
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="rowCount">要添加的行数</param>
        /// <param name="app">Word应用程序对象（用于更新状态栏和Selection）</param>
        public static void BatchAddRows(Table tbl, int rowCount, Application app = null)
        {
            if (tbl == null || rowCount <= 0) return;

            // 如果app为null，回退到逐行添加
            if (app == null)
            {
                try
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        tbl.Rows.Add();
                    }
                }
                catch
                {
                    // 忽略错误
                }
                return;
            }

            try
            {
                const int BATCH_SIZE = 200; // 每批插入200行
                int remaining = rowCount;
                int added = 0;

                while (remaining > 0)
                {
                    int batch = System.Math.Min(remaining, BATCH_SIZE);

                    try
                    {
                        // 选中表格最后一行，使用 InsertRowsBelow 批量插入
                        tbl.Rows[tbl.Rows.Count].Select();
                        app.Selection.InsertRowsBelow(batch);
                    }
                    catch
                    {
                        // 如果批量插入失败，回退到逐行添加
                        for (int i = 0; i < batch; i++)
                        {
                            try { tbl.Rows.Add(); } catch (Exception rowEx) { System.Diagnostics.Debug.WriteLine("逐行添加表格行失败: " + rowEx.Message); break; }
                        }
                    }

                    added += batch;
                    remaining -= batch;

                    // 预分配行数进度由进度窗口统一显示，不再更新状态栏
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 预分配表格行数
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="estimatedImageCount">预计图片数量</param>
        /// <param name="imagesPerRow">每行图片数</param>
        /// <param name="needDescription">是否需要描述行</param>
        /// <param name="app">Word应用程序对象</param>
        public static void PreAllocateRows(Table tbl, int estimatedImageCount, 
            int imagesPerRow = 2, bool needDescription = false, Application app = null)
        {
            if (tbl == null) return;

            try
            {
                // 安全检查
                if (imagesPerRow <= 0) imagesPerRow = 2;

                // 计算需要的行数
                int neededRows = (estimatedImageCount + imagesPerRow - 1) / imagesPerRow;

                // 需要描述行时加倍
                if (needDescription)
                {
                    neededRows *= 2;
                }

                // 限制最大预分配行数
                const int MAX_PREALLOCATE_ROWS = 1000;
                if (neededRows > MAX_PREALLOCATE_ROWS)
                {
                    neededRows = MAX_PREALLOCATE_ROWS;
                }

                // 批量添加行
                BatchAddRows(tbl, neededRows, app);
            }
            catch
            {
                // 忽略错误
            }
        }

        #endregion
    }
}
