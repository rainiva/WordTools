using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordTools.Services
{
    /// <summary>
    /// 表格工具服务
    /// 处理表格相关的操作、验证和自动编号
    /// </summary>
    public static class TableService
    {
        private static void SafeIgnore(Exception ex, string context)
        {
            Debug.WriteLine($"{context}: {ex.Message}");
        }

        #region 表格验证

        /// <summary>
        /// 验证当前选择是否在表格中
        /// </summary>
        public static bool IsSelectionInTable(Selection selection)
        {
            if (selection == null) return false;
            return (bool)selection.Information[WdInformation.wdWithInTable];
        }

        /// <summary>
        /// 验证当前选择是否在第一列
        /// </summary>
        public static bool IsSelectionInFirstColumn(Selection selection)
        {
            if (!IsSelectionInTable(selection)) return false;
            return selection.Cells[1].ColumnIndex == 1;
        }

        /// <summary>
        /// 获取当前表格对象
        /// </summary>
        public static Table GetCurrentTable(Selection selection)
        {
            if (!IsSelectionInTable(selection)) return null;
            return selection.Tables[1];
        }

        /// <summary>
        /// 检查单元格是否适合插入图片
        /// </summary>
        public static bool IsCellSuitableForImage(Cell targetCell)
        {
            return ImageRowPlanner.CanHostSingleImage(GetCellAvailability(targetCell, null));
        }

        public static ImageCellAvailability GetCellAvailability(Cell targetCell)
        {
            return GetCellAvailability(targetCell, null);
        }

        public static ImageCellAvailability GetCellAvailability(Cell targetCell, ImageInsertionBatchContext context)
        {
            if (targetCell == null) return ImageCellAvailability.Blocked;

            Stopwatch availabilityWatch = Stopwatch.StartNew();

            try
            {
                // 检查合并单元格（合并单元格的索引访问可能异常）
                try
                {
                    int rowIdx = targetCell.RowIndex;
                    int colIdx = targetCell.ColumnIndex;
                    if (rowIdx < 1 || colIdx < 1)
                        return ImageCellAvailability.Blocked;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[TableService] GetImageCellAvailability index access error: {ex.Message}");
                    return ImageCellAvailability.Blocked;
                }

                // 检查是否已有内联图片
                if (targetCell.Range.InlineShapes.Count > 0)
                {
                    return ImageCellAvailability.OverwriteImage;
                }

                if (HasFloatingShapeInCell(targetCell, context))
                {
                    return ImageCellAvailability.OverwriteImage;
                }

                // 检查单元格是否包含编号（SEQ域或纯文本格式）
                bool hasNumbering = HasNumbering(targetCell);

                // 如果包含编号，清除编号使其可复用
                if (hasNumbering)
                {
                    try
                    {
                        // 删除所有 SEQ 域（向后兼容）
                        Range r = targetCell.Range;
                        r.SetRange(r.Start, r.End - 1);
                        for (int i = r.Fields.Count; i >= 1; i--)
                        {
                            try
                            {
                                if (r.Fields[i].Type == WdFieldType.wdFieldSequence)
                                    r.Fields[i].Delete();
                            }
                            catch (Exception ex) { SafeIgnore(ex, "表格操作失败"); }
                        }
                        // 重新获取 Range 并清除剩余文本
                        r = targetCell.Range;
                        r.SetRange(r.Start, r.End - 1);
                        r.Text = "";
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[TableService] GetImageCellAvailability clear numbering error: {ex.Message}");
                        return ImageCellAvailability.Blocked;
                    }
                    return ImageCellAvailability.Available;
                }

                // 获取单元格文本并清理
                string cellText = CleanCellText(targetCell.Range.Text);

                // 空单元格适合插入
                if (string.IsNullOrEmpty(cellText))
                {
                    return ImageCellAvailability.Available;
                }

                // 检查是否为文本编号格式（如 "1.", "2." 等）
                // 如果是编号，清除文本使其可复用
                if (TableNumberingService.ExtractNumberFromCellText(cellText).HasValue)
                {
                    // 清除序号文本，使单元格可以被复用插入图片
                    try
                    {
                        targetCell.Range.Text = "";
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[TableService] GetImageCellAvailability clear number text error: {ex.Message}");
                        return ImageCellAvailability.Blocked;
                    }
                    return ImageCellAvailability.Available;
                }

                return ImageCellAvailability.OverwriteText;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] GetImageCellAvailability error: {ex.Message}");
                return ImageCellAvailability.Blocked;
            }
            finally
            {
                availabilityWatch.Stop();
                if (context != null)
                {
                    context.Diagnostics.RecordCellAvailabilityCheck(availabilityWatch.ElapsedMilliseconds);
                }
            }
        }

        /// <summary>
        /// 找到下一个适合插入图片的单元格
        /// </summary>
        public static bool FindNextSuitableCell(Table tbl, int startRow, 
            out int foundRow, out int foundCol, int preferredCol = 1)
        {
            foundRow = startRow;
            foundCol = preferredCol;

            if (tbl == null) return false;

            try
            {
                int maxRow = startRow + 10; // 最多搜索10行

                for (int row = startRow; row <= maxRow; row++)
                {
                    EnsureRowExists(tbl, row);

                    // 检查列数是否足够（防止列数中途变化）
                    if (tbl.Columns.Count < 1)
                        continue;

                    // 检查第1列（带合并单元格保护）
                    bool col1Suitable = false;
                    try
                    {
                        col1Suitable = IsCellSuitableForImage(tbl.Cell(row, 1));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[TableService] FindNextAvailableImageCell col1 check error: {ex.Message}");
                        col1Suitable = false;
                    }

                    // 如果第1列不适合，跳过整行
                    if (!col1Suitable) continue;

                    // 检查列数是否足够访问第2列
                    bool col2Suitable = false;
                    if (tbl.Columns.Count >= 2)
                    {
                        try
                        {
                            col2Suitable = IsCellSuitableForImage(tbl.Cell(row, 2));
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[TableService] FindNextAvailableImageCell col2 check error: {ex.Message}");
                            col2Suitable = false;
                        }
                    }

                    // 根据优先列返回结果
                    if (preferredCol == 2)
                    {
                        // 优先第2列：先检查第2列，再检查第1列
                        if (col2Suitable)
                        {
                            foundRow = row;
                            foundCol = 2;
                            return true;
                        }
                        if (col1Suitable)
                        {
                            foundRow = row;
                            foundCol = 1;
                            return true;
                        }
                    }
                    else
                    {
                        // 优先第1列：先检查第1列，再检查第2列
                        if (col1Suitable)
                        {
                            foundRow = row;
                            foundCol = 1;
                            return true;
                        }
                        if (col2Suitable)
                        {
                            foundRow = row;
                            foundCol = 2;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] FindNextAvailableImageCell error: {ex.Message}");
            }

            return false;
        }

        private static bool HasFloatingShapeInCell(Cell targetCell, ImageInsertionBatchContext context)
        {
            Stopwatch lookupWatch = Stopwatch.StartNew();
            try
            {
                Range cellRange = targetCell.Range;
                FloatingShapeIndex index = context != null
                    ? context.GetOrCreateFloatingShapeIndex(() => BuildFloatingShapeIndex(cellRange.Document.Shapes))
                    : BuildFloatingShapeIndex(cellRange.Document.Shapes);

                return index.HasShapeInRange(cellRange.Start, cellRange.End);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] HasFloatingShapeInCell error: {ex.Message}");
            }
            finally
            {
                lookupWatch.Stop();
                if (context != null)
                {
                    context.Diagnostics.RecordFloatingShapeLookup(lookupWatch.ElapsedMilliseconds);
                }
            }

            return false;
        }

        private static FloatingShapeIndex BuildFloatingShapeIndex(Shapes shapes)
        {
            var anchors = new List<FloatingShapeAnchor>();
            if (shapes == null)
            {
                return FloatingShapeIndex.Create(anchors);
            }

            for (int i = 1; i <= shapes.Count; i++)
            {
                try
                {
                    var shape = shapes[i];
                    if (shape != null && shape.Anchor != null)
                    {
                        anchors.Add(new FloatingShapeAnchor(shape.Anchor.Start, shape.Anchor.End));
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[TableService] BuildFloatingShapeIndex iteration error: {ex.Message}");
                }
            }

            return FloatingShapeIndex.Create(anchors);
        }

        #endregion

        #region 合并单元格检测

        /// <summary>
        /// 检测指定坐标是否落在合并单元格区域内
        /// 原理：Word COM 中，合并区域内部任意位置调用 tbl.Cell() 都会返回左上角同一个 Cell 对象
        /// </summary>
        public static bool IsMergedCell(Table tbl, int row, int col, out int mergeTopRow, out int mergeLeftCol)
        {
            mergeTopRow = row;
            mergeLeftCol = col;
            if (tbl == null || row < 1 || col < 1) return false;

            try
            {
                Cell cell = tbl.Cell(row, col);
                int actualRow = cell.RowIndex;
                int actualCol = cell.ColumnIndex;

                // 如果返回的坐标与请求坐标不一致，说明落在合并区域内
                if (actualRow != row || actualCol != col)
                {
                    mergeTopRow = actualRow;
                    mergeLeftCol = actualCol;
                    return true;
                }

                // 额外检测：横向合并（通过 Cells.Count）
                if (tbl.Rows.Count >= row && tbl.Columns.Count > 1)
                {
                    if (ShouldTreatCellCountMismatchAsMerged(col, tbl.Rows[row].Cells.Count, tbl.Columns.Count))
                    {
                        mergeTopRow = row;
                        mergeLeftCol = col;
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] IsMergedCell error: {ex.Message}");
                mergeTopRow = row;
                mergeLeftCol = col;
                return true;
            }
        }

        /// <summary>
        /// 获取合并单元格占用的总行数（纵向跨度）
        /// </summary>
        public static int GetMergedRowSpan(Table tbl, int topRow, int leftCol)
        {
            try
            {
                int span = 1;
                while (topRow + span <= tbl.Rows.Count)
                {
                    Cell nextCell = tbl.Cell(topRow + span, leftCol);
                    if (nextCell.RowIndex == topRow && nextCell.ColumnIndex == leftCol)
                        span++;
                    else
                        break;
                }
                return span;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] GetMergedRowSpan error: {ex.Message}");
                return 1;
            }
        }

        public static ImageCellAvailability GetCellAvailability(Table tbl, int row, int col)
        {
            return GetCellAvailability(tbl, row, col, null);
        }

        public static ImageCellAvailability GetCellAvailability(Table tbl, int row, int col, ImageInsertionBatchContext context)
        {
            if (tbl == null || row < 1 || col < 1)
            {
                return ImageCellAvailability.Blocked;
            }

            if (context != null && context.TryGetCachedRowAvailability(row, out var cachedRow))
            {
                return col == 1 ? cachedRow.LeftCell : cachedRow.RightCell;
            }

            try
            {
                ImageRowAvailability rowAvailability = GetImageRowAvailability(tbl, row, context);
                return col == 1 ? rowAvailability.LeftCell : rowAvailability.RightCell;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] GetCellAvailability error: {ex.Message}");
                return ImageCellAvailability.Blocked;
            }
        }

        public static ImageRowAvailability GetImageRowAvailability(Table tbl, int row)
        {
            return GetImageRowAvailability(tbl, row, null);
        }

        public static ImageRowAvailability GetImageRowAvailability(Table tbl, int row, ImageInsertionBatchContext context)
        {
            if (context != null && context.TryGetCachedRowAvailability(row, out var cachedRow))
            {
                return cachedRow;
            }

            ImageRowAvailability rowAvailability = new ImageRowAvailability(
                row,
                GetCellAvailabilityCore(tbl, row, 1, context),
                GetCellAvailabilityCore(tbl, row, 2, context));

            if (context != null)
            {
                context.CacheRowAvailability(rowAvailability);
            }

            return rowAvailability;
        }

        public static bool FindNextSuitableImageRow(Table tbl, int startRow, out int foundRow, List<int> mergedCellRows = null)
        {
            return FindNextSuitableImageRow(tbl, startRow, out foundRow, mergedCellRows, null);
        }

        public static bool FindNextSuitableImageRow(Table tbl, int startRow, out int foundRow, List<int> mergedCellRows, ImageInsertionBatchContext context)
        {
            foundRow = startRow;

            if (tbl == null)
            {
                return false;
            }

            var currentRow = GetImageRowAvailability(tbl, startRow, context);
            if (currentRow.HasMergedCell)
            {
                AddMergedRow(mergedCellRows, currentRow.RowIndex);
            }

            int maxRow = GetImageRowSearchEndRow(startRow, tbl.Rows.Count);
            int preferredRow = ImageRowPlanner.FindPreferredPairRow(
                currentRow,
                EnumerateFallbackRows(tbl, startRow + 1, maxRow, mergedCellRows, context),
                -1);

            if (preferredRow < 0)
            {
                return false;
            }

            foundRow = preferredRow;
            return true;
        }

        #endregion

        #region 表格操作

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

        public static bool ShouldTreatCellCountMismatchAsMerged(int requestedColumn, int visibleCellCount, int totalColumnCount)
        {
            return requestedColumn > visibleCellCount && visibleCellCount < totalColumnCount;
        }

        public static int GetImageRowSearchEndRow(int startRow, int lastExistingRow)
        {
            return lastExistingRow > startRow ? lastExistingRow : startRow;
        }

        private static IEnumerable<ImageRowAvailability> EnumerateFallbackRows(
            Table tbl,
            int startRow,
            int endRow,
            List<int> mergedCellRows,
            ImageInsertionBatchContext context)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                EnsureRowExists(tbl, row);

                var rowState = GetImageRowAvailability(tbl, row, context);
                if (rowState.HasMergedCell)
                {
                    AddMergedRow(mergedCellRows, rowState.RowIndex);
                }

                yield return rowState;
            }
        }

        private static ImageCellAvailability GetCellAvailabilityCore(Table tbl, int row, int col, ImageInsertionBatchContext context)
        {
            EnsureRowExists(tbl, row);

            if (tbl.Columns.Count < col)
            {
                return ImageCellAvailability.Blocked;
            }

            int mergeTopRow;
            int mergeLeftCol;
            if (IsMergedCell(tbl, row, col, out mergeTopRow, out mergeLeftCol))
            {
                return ImageCellAvailability.Merged;
            }

            try
            {
                return GetCellAvailability(tbl.Cell(row, col), context);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] GetCellAvailabilityCore error: {ex.Message}");
                return ImageCellAvailability.Merged;
            }
        }

        /// <summary>
        /// 确保表格行存在（轻量版，不调用 AdjustTableColumns）
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="rowIndex">需要的行索引</param>
        /// <param name="cachedRowCount">缓存的行数（ref，调用后更新为实际行数）</param>
        public static void EnsureRowExists(Table tbl, int rowIndex, ref int cachedRowCount)
        {
            if (tbl == null) return;

            try
            {
                int rowsToAdd = rowIndex - cachedRowCount;

                if (rowsToAdd > 0)
                {
                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        tbl.Rows.Add();
                    }
                    // 更新缓存值
                    cachedRowCount = tbl.Rows.Count;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] EnsureRowExists (cached) error: {ex.Message}");
            }
        }

        /// <summary>
        /// 确保表格行存在（兼容旧版，内部查询行数）
        /// </summary>
        public static void EnsureRowExists(Table tbl, int rowIndex)
        {
            if (tbl == null) return;

            try
            {
                int currentRowCount = tbl.Rows.Count;
                int rowsToAdd = rowIndex - currentRowCount;

                if (rowsToAdd > 0)
                {
                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        tbl.Rows.Add();
                    }
                }

                // 确保有2列
                AdjustTableColumns(tbl, 2);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] EnsureRowExists error: {ex.Message}");
            }
        }

        /// <summary>
        /// 调整表格列数
        /// </summary>
        public static void AdjustTableColumns(Table tbl, int targetColCount)
        {
            if (tbl == null) return;

            try
            {
                int currentColCount = tbl.Columns.Count;

                if (currentColCount > targetColCount)
                {
                    // 删除多余的列
                    for (int i = currentColCount; i > targetColCount; i--)
                    {
                        tbl.Columns[i].Delete();
                    }
                }
                else if (currentColCount < targetColCount)
                {
                    // 添加缺少的列
                    for (int i = 0; i < targetColCount - currentColCount; i++)
                    {
                        tbl.Columns.Add();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] AdjustTableColumns error: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查表格是否使用固定列宽
        /// </summary>
        public static bool IsTableFixedColumnWidth(Table tbl)
        {
            if (tbl == null) return false;
            try
            {
                return !tbl.AllowAutoFit;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] IsTableFixedColumnWidth error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 设置表格为固定列宽
        /// </summary>
        public static void SetTableFixedColumnWidth(Table tbl)
        {
            if (tbl == null) return;
            try
            {
                tbl.AllowAutoFit = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] SetTableFixedColumnWidth error: {ex.Message}");
            }
        }

        #endregion

        #region 标题行和描述行

        /// <summary>
        /// 创建标题行
        /// </summary>
        public static void CreateTitleRow(Table tbl, ref int rowIndex, string titleText)
        {
            if (tbl == null) return;

            try
            {
                AdjustTableColumns(tbl, 2);
                EnsureRowExists(tbl, rowIndex);

                // 检查当前行是否为空
                bool currentRowIsEmpty = false;
                try
                {
                    if (tbl.Cell(rowIndex, 1).Range.InlineShapes.Count == 0 &&
                        tbl.Cell(rowIndex, 2).Range.InlineShapes.Count == 0)
                    {
                        string cell1Text = CleanCellText(tbl.Cell(rowIndex, 1).Range.Text);
                        string cell2Text = CleanCellText(tbl.Cell(rowIndex, 2).Range.Text);
                        if (string.IsNullOrEmpty(cell1Text) && string.IsNullOrEmpty(cell2Text))
                        {
                            currentRowIsEmpty = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[TableService] CreateTitleRow empty check error: {ex.Message}");
                }

                // 如果当前行不为空，插入新行
                if (!currentRowIsEmpty)
                {
                    rowIndex++;
                    tbl.Rows.Add();
                }

                AdjustTableColumns(tbl, 2);
                tbl.Rows.Add(); // 添加下一行用于内容

                // 合并单元格作为标题行
                tbl.Cell(rowIndex, 1).Merge(tbl.Cell(rowIndex, 2));

                // 写入标题
                tbl.Cell(rowIndex, 1).Range.Text = titleText;
                tbl.Rows[rowIndex].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 移动到下一行
                rowIndex++;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] CreateTitleRow error: {ex.Message}");
            }
        }

        /// <summary>
        /// 插入描述行
        /// </summary>
        public static void InsertDescriptionRow(Table tbl, ref int rowIndex)
        {
            if (tbl == null) return;

            EnsureRowExists(tbl, rowIndex);
            AdjustTableColumns(tbl, 2);
        }

        /// <summary>
        /// 插入文件名描述行
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="rowIndex">行索引（引用，方法内会递增）</param>
        /// <param name="descriptions">描述文本数组（可以是文件路径或纯文本）</param>
        /// <param name="isFilePath">描述是否为文件路径（true=提取文件名，false=直接使用）</param>
        public static void InsertFileNameDescriptionRow(Table tbl, ref int rowIndex, string[] descriptions, bool isFilePath = true)
        {
            if (tbl == null) return;

            try
            {
                EnsureRowExists(tbl, rowIndex);
                AdjustTableColumns(tbl, 2);

                // 插入描述文本到对应列
                for (int i = 0; i < Math.Min(descriptions.Length, 2); i++)
                {
                    string displayText = isFilePath
                        ? FileService.GetFileNameWithoutExtension(descriptions[i])
                        : descriptions[i];
                    var cell = tbl.Cell(rowIndex, i + 1);
                    cell.Range.Text = displayText;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                // 如果只有一个描述，第二列留空
                if (descriptions.Length < 2)
                {
                    var cell2 = tbl.Cell(rowIndex, 2);
                    cell2.Range.Text = "";
                    cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    cell2.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                rowIndex++;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] InsertFileNameDescriptionRow error: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充空单元格为 N/A
        /// </summary>
        public static void FillEmptyCellsWithNA(Table tbl, int rowIndex, int startCol, int endCol)
        {
            if (tbl == null) return;

            try
            {
                if (rowIndex < 1 || rowIndex > tbl.Rows.Count) return;
                if (startCol < 1 || endCol > tbl.Columns.Count) return;

                for (int col = startCol; col <= endCol; col++)
                {
                    var cell = tbl.Cell(rowIndex, col);
                    cell.Range.Text = "N/A";
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TableService] FillEmptyCellsWithNA error: {ex.Message}");
            }
        }

        #endregion

        /// <summary>
        /// 检查单元格是否有编号（SEQ域或纯文本格式）
        /// </summary>
        internal static bool HasNumbering(Cell cell)
        {
            try
            {
                foreach (Field f in cell.Range.Fields)
                {
                    if (f.Type == WdFieldType.wdFieldSequence)
                        return true;
                }
                Range r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                string text = r.Text ?? "";
                text = text.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                return Regex.IsMatch(text, @"^\d+\.");
            }
            catch (Exception ex) { SafeIgnore(ex, "表格操作失败"); }
            return false;
        }

        #region 辅助方法

        /// <summary>
        /// 清理单元格文本
        /// </summary>
        internal static string CleanCellText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("\r", "").Replace("\n", "")
                .Replace("\t", "").Replace("\a", "").Replace("\u00A0", "").Trim();
        }

        #endregion
    }
}
