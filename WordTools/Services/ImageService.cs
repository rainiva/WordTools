using System;
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

        #region 图片插入

        /// <summary>
        /// 插入图片到单元格并调整尺寸
        /// </summary>
        /// <param name="targetCell">目标单元格</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="minHeightPoints">最小高度（磅），-1 表示不限制</param>
        /// <returns>插入的图片对象</returns>
        public static InlineShape InsertImageToCell(Cell targetCell, string imagePath, float minHeightPoints = -1)
        {
            if (targetCell == null || string.IsNullOrEmpty(imagePath))
                return null;

            try
            {
                // 获取单元格的宽度和高度
                float cellWidth = targetCell.Width;
                float cellHeight = targetCell.Height;

                // 计算目标尺寸（考虑边距）
                float targetWidth = cellWidth - 6; // 减去左右边距各3磅
                float targetHeight = cellHeight - 6; // 减去上下边距各3磅

                if (targetWidth < 10) targetWidth = cellWidth;
                if (targetHeight < 10) targetHeight = cellHeight;

                // 插入图片到单元格范围
                var p = targetCell.Range.InlineShapes.AddPicture(
                    FileName: imagePath,
                    LinkToFile: false,
                    SaveWithDocument: true);

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

                // 应用缩放
                if (scaleRatio < 1)
                {
                    p.Width = p.Width * scaleRatio;
                    p.Height = p.Height * scaleRatio;
                }

                // 如果设置了最小高度且当前高度小于最小高度，调整高度
                if (minHeightPoints > 0 && p.Height < minHeightPoints)
                {
                    p.Height = minHeightPoints;
                }

                return p;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 快速插入图片（最小化内存使用）
        /// </summary>
        /// <param name="targetCell">目标单元格</param>
        /// <param name="imagePath">图片文件路径</param>
        /// <param name="minHeightPoints">最小高度（磅），-1 表示不限制</param>
        public static void InsertImageFast(Cell targetCell, string imagePath, float minHeightPoints = -1)
        {
            if (targetCell == null || string.IsNullOrEmpty(imagePath))
                return;

            try
            {
                // 获取单元格宽度
                float cellWidth = targetCell.Width - 6;

                // 清空单元格现有内容
                targetCell.Range.Text = "";

                // 插入图片
                var p = targetCell.Range.InlineShapes.AddPicture(
                    FileName: imagePath,
                    LinkToFile: false,
                    SaveWithDocument: true);

                // msoTrue = -1
                ((dynamic)p).LockAspectRatio = -1;

                // 快速调整：只检查宽度限制
                if (p.Width > cellWidth)
                {
                    p.Width = cellWidth;
                }

                // 最小高度限制
                if (minHeightPoints > 0 && p.Height < minHeightPoints)
                {
                    p.Height = minHeightPoints;
                }
            }
            catch (Exception)
            {
                // 忽略单个图片插入错误
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
                                    if (shape.Width > cell.Width - 6)
                                    {
                                        shape.Width = cell.Width - 6;
                                    }

                                    if (minHeightPoints > 0 && shape.Height < minHeightPoints)
                                    {
                                        shape.Height = minHeightPoints;
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
        /// <param name="app">Word应用程序对象（用于更新状态栏）</param>
        public static void BatchAddRows(Table tbl, int rowCount, Application app = null)
        {
            if (tbl == null || rowCount <= 0) return;

            try
            {
                const int BATCH_SIZE = 100;
                int batchCount = 0;

                for (int i = 1; i <= rowCount; i++)
                {
                    tbl.Rows.Add();
                    batchCount++;

                    // 每100行更新一次状态栏
                    if (batchCount >= BATCH_SIZE)
                    {
                        if (app != null)
                        {
                            app.StatusBar = string.Format("正在准备表格... {0}/{1}", i, rowCount);
                        }
                        System.Windows.Forms.Application.DoEvents();
                        batchCount = 0;
                    }
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
