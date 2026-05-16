using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                catch
                {
                    // 合并单元格通常无法正确获取索引
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
                            catch { }
                        }
                        // 重新获取 Range 并清除剩余文本
                        r = targetCell.Range;
                        r.SetRange(r.Start, r.End - 1);
                        r.Text = "";
                    }
                    catch
                    {
                        // 清除失败则不适合插入
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
                if (ExtractNumberFromCellText(cellText).HasValue)
                {
                    // 清除序号文本，使单元格可以被复用插入图片
                    try
                    {
                        targetCell.Range.Text = "";
                    }
                    catch
                    {
                        // 清除失败则不适合插入
                        return ImageCellAvailability.Blocked;
                    }
                    return ImageCellAvailability.Available;
                }

                return ImageCellAvailability.OverwriteText;
            }
            catch
            {
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
                    catch
                    {
                        // 合并单元格或索引越界时跳过
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
                        catch
                        {
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
            catch
            {
                // 忽略错误
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
            catch
            {
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
                catch
                {
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
            catch
            {
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
            catch
            {
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
            catch
            {
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
            catch
            {
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
            catch
            {
                // 忽略错误
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
            catch
            {
                // 忽略错误
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
            catch
            {
                // 忽略错误
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
            catch
            {
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
            catch
            {
                // 忽略错误
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
                catch
                {
                    // 忽略错误
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
            catch
            {
                // 忽略错误
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
            catch
            {
                // 忽略错误
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
            catch
            {
                // 忽略错误
            }
        }

        #endregion

        #region 自动编号

        /// <summary>
        /// 刷新表格编号（从光标处开始检查更新）
        /// 自动检测光标位置，只处理光标行及之后的编号
        /// </summary>
        public static void RefreshTableNumbering(Table tbl, Document doc, int alignment = 2, 
            Action<string> progressCallback = null)
        {
            if (tbl == null || doc == null) return;
                
            Application app = null;
            bool wasScreenUpdating = true;
                
            try
            {
                try
                {
                    app = tbl.Range.Application;
                    wasScreenUpdating = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                }
                catch { }
                
                WdParagraphAlignment wdAlignment;
                switch (alignment)
                {
                    case 1: wdAlignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                    case 3: wdAlignment = WdParagraphAlignment.wdAlignParagraphRight; break;
                    default: wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter; break;
                }
                
                int colCount = tbl.Columns.Count;
                int totalRows = tbl.Rows.Count;
                DateTime startTime = DateTime.Now;
                
                // === Step 1: 定位光标行，确定扫描范围 ===
                int cursorRow = 1;
                try { cursorRow = app.Selection.Range.Cells[1].RowIndex; }
                catch { cursorRow = 1; }
                
                // 扫描从光标前一行开始（需判断"图片行→描述行"配对）
                int scanStart = Math.Max(1, cursorRow - 1);
                
                // 临时启用ScreenUpdating让状态栏可见
                try { if (app != null) app.ScreenUpdating = true; } catch { }
                progressCallback?.Invoke(string.Format("从第 {0} 行开始检查编号...", scanStart));
                System.Windows.Forms.Application.DoEvents();
                try { if (app != null) app.ScreenUpdating = false; } catch { }
        
                // === 快速路径尝试：跳过 InlineShapes 扫描，直接扫描编号行 ===
                // 阶段 A：直接扫描编号行（替代 InlineShapes + imageRow 方式）
                var numberedRows = new List<int>(); // 有编号的行
                var rowNumbers = new Dictionary<int, Dictionary<int, int>>(); // row -> col -> currentNumber
        
                for (int row = scanStart; row <= totalRows; row++)
                {
                    try
                    {
                        Range r = tbl.Cell(row, 1).Range;
                        r.SetRange(r.Start, r.End - 1);
                        string text = (r.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        var m = Regex.Match(text, @"^(\d+)\.");
                        if (m.Success)
                        {
                            numberedRows.Add(row);
                            var colNums = new Dictionary<int, int>();
                            colNums[1] = int.Parse(m.Groups[1].Value);
                            // 检查其他列
                            for (int col = 2; col <= colCount; col++)
                            {
                                try
                                {
                                    Range r2 = tbl.Cell(row, col).Range;
                                    r2.SetRange(r2.Start, r2.End - 1);
                                    string t2 = (r2.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                                    var m2 = Regex.Match(t2, @"^(\d+)\.");
                                    if (m2.Success) colNums[col] = int.Parse(m2.Groups[1].Value);
                                }
                                catch { }
                            }
                            rowNumbers[row] = colNums;
                        }
                    }
                    catch { }
                            
                    // 每50行做一次DoEvents
                    if ((row - scanStart) % 50 == 0)
                        System.Windows.Forms.Application.DoEvents();
                }
        
                // 如果快速路径找到了编号行，尝试快速修复
                if (numberedRows.Count > 0)
                {
                    // 阶段 B：计算 baseNumber（简化版，不验证图片行）
                    int baseNumber = 0;
                    if (scanStart > 1)
                    {
                        for (int row = scanStart - 1; row >= 1; row--)
                        {
                            try
                            {
                                int? num = ExtractNumberFromCell(tbl.Cell(row, colCount));
                                if (num.HasValue)
                                {
                                    baseNumber = num.Value;
                                    break;
                                }
                            }
                            catch { }
                        }
                    }
        
                    // 阶段 C：检查连续性（从缓存，零 COM 调用）
                    int expectedNum = baseNumber;
                    int firstBadIdx = -1;
                    bool allContinuous = true;
                            
                    for (int i = 0; i < numberedRows.Count; i++)
                    {
                        int row = numberedRows[i];
                        for (int col = 1; col <= colCount; col++)
                        {
                            expectedNum++;
                            int actual = rowNumbers.ContainsKey(row) && rowNumbers[row].ContainsKey(col)
                                ? rowNumbers[row][col] : -1;
                            if (actual != expectedNum)
                            {
                                firstBadIdx = i;
                                allContinuous = false;
                                goto checkDone;
                            }
                        }
                    }
                    checkDone:
        
                    // 如果编号只是值不对（不连续），使用快速路径直接更新
                    if (!allContinuous && firstBadIdx >= 0)
                    {
                        // 阶段 D：快速更新（跳过正确的单元格）
                        int currentNum = baseNumber + firstBadIdx * colCount;
                        int cellsUpdated = 0;
                        int totalCellsToUpdate = (numberedRows.Count - firstBadIdx) * colCount;
        
                        // 显示进度
                        try { if (app != null) app.ScreenUpdating = true; } catch { }
                        progressCallback?.Invoke("正在快速更新编号...");
                        System.Windows.Forms.Application.DoEvents();
                        try { if (app != null) app.ScreenUpdating = false; } catch { }
        
                        for (int i = firstBadIdx; i < numberedRows.Count; i++)
                        {
                            int row = numberedRows[i];
                            for (int col = 1; col <= colCount; col++)
                            {
                                currentNum++;
                                // 跳过已正确的
                                if (rowNumbers.ContainsKey(row) && rowNumbers[row].ContainsKey(col) && 
                                    rowNumbers[row][col] == currentNum)
                                    continue;
                                try { UpdateCellNumber(tbl.Cell(row, col), currentNum); }
                                catch { }
                                cellsUpdated++;
                            }
                                    
                            // 每50行做一次DoEvents
                            if ((i - firstBadIdx) % 50 == 0)
                                System.Windows.Forms.Application.DoEvents();
                        }
        
                        progressCallback?.Invoke(string.Format("编号刷新完成！(快速路径, 更新{0}个编号, 耗时{1:F2}s)",
                            cellsUpdated, (DateTime.Now - startTime).TotalSeconds));
                        return;
                    }
        
                    // 如果所有编号都正确，直接返回
                    if (allContinuous)
                    {
                        progressCallback?.Invoke(string.Format("编号已是最新，无需更新 (快速路径, 耗时 {0:F2}s)", 
                            (DateTime.Now - startTime).TotalSeconds));
                        return;
                    }
                }
        
                // === 快速路径失败或没有找到编号行，回退到完整流程（包含 InlineShapes 扫描） ===
                        
                // === Step 2: 扫描描述行编号（基于图片行位置识别描述行，提取现有编号值） ===
                // 优化：全列缓存，消除后续阶段的重复COM调用
                var numbersByRowCol = new Dictionary<int, Dictionary<int, int>>(); // row -> (col -> number)
                var descriptionRows = new HashSet<int>(); // 所有描述行行号
                var numbersByRow = new Dictionary<int, int>(); // 第1列的快速查找缓存
                        
                // 先扫描 InlineShapes 确定图片行位置
                var imageRows = new HashSet<int>();
                int shapeCount = 0;
                try
                {
                    Range shapeScanRange = doc.Range(tbl.Cell(scanStart, 1).Range.Start, tbl.Range.End);
                    foreach (InlineShape shape in shapeScanRange.InlineShapes)
                    {
                        try { imageRows.Add(shape.Range.Cells[1].RowIndex); }
                        catch { }
                        // 每50个图片让出UI线程
                        shapeCount++;
                        if (shapeCount % 50 == 0)
                            System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch { }
                        
                // 基于图片行识别描述行并提取所有列的编号值（全列缓存）
                var sortedImageRows = new List<int>(imageRows);
                sortedImageRows.Sort();
                foreach (int imgRow in sortedImageRows)
                {
                    int descRow = imgRow + 1;
                    if (descRow <= totalRows && !imageRows.Contains(descRow))
                    {
                        descriptionRows.Add(descRow);
                        var colNumbers = new Dictionary<int, int>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            try
                            {
                                int? num = ExtractNumberFromCell(tbl.Cell(descRow, col));
                                if (num.HasValue)
                                    colNumbers[col] = num.Value;
                            }
                            catch { }
                        }
                        if (colNumbers.Count > 0)
                        {
                            numbersByRowCol[descRow] = colNumbers;
                            // 保留第1列值作为快速查找
                            if (colNumbers.ContainsKey(1))
                                numbersByRow[descRow] = colNumbers[1];
                        }
                    }
                }
        
                // === 计算 scanStart 之前的编号基数 ===
                // 当从表格中间刷新时，需要接续之前的编号，而不是从1开始
                // 关键：必须验证找到的行是真正的描述行（上方有图片），避免命中孤立行
                int baseNumberSlow = 0;
                if (scanStart > 1)
                {
                    // 从 scanStart-1 往前找最近的有编号的行
                    for (int row = scanStart - 1; row >= 1; row--)
                    {
                        try
                        {
                            // 检查最后一列是否有编号
                            int? num = ExtractNumberFromCell(tbl.Cell(row, colCount));
                            if (num.HasValue)
                            {
                                // 验证：真正的描述行必须前一行有图片
                                // 先检查 imageRows 缓存（已扫描的范围）
                                bool prevRowHasImage = imageRows.Contains(row - 1);
                                        
                                // 如果不在缓存中（可能是 scanStart 之前的行），额外检查
                                if (!prevRowHasImage && row - 1 >= 1)
                                {
                                    try
                                    {
                                        Cell prevCell = tbl.Cell(row - 1, 1);
                                        prevRowHasImage = prevCell.Range.InlineShapes.Count > 0;
                                    }
                                    catch { }
                                }
                                        
                                // 只有前一行有图片，才是真正的描述行，才能作为 baseNumber
                                if (prevRowHasImage)
                                {
                                    baseNumberSlow = num.Value;
                                    break;
                                }
                                // 如果不是真正的描述行（孤立行），继续往前找
                            }
                        }
                        catch { }
                    }
                }
        
                // === 快速路径：检查编号值连续性，不连续则直接批量更新 ===
                // 最常见场景：删行后值不对
                // 注意：多列表格中，每个描述行的各列应该有独立递增的编号
                // 例如2列表格：描述行1 col1=1, col2=2; 描述行2 col1=3, col2=4
                var sortedRows = new List<int>(numbersByRow.Keys);
                sortedRows.Sort();
        
                int firstBadIdxSlow = -1;
                int expectedNumSlow = baseNumberSlow; // 期望的下一个编号值（从baseNumber开始）
                for (int i = 0; i < sortedRows.Count; i++)
                {
                    int row = sortedRows[i];
                    bool rowOk = true;
                            
                    // 检查该行的所有列是否连续递增（使用缓存，零COM调用）
                    for (int col = 1; col <= colCount; col++)
                    {
                        expectedNumSlow++;
                        int? actualNum = null;
                        if (numbersByRowCol.ContainsKey(row) && numbersByRowCol[row].ContainsKey(col))
                            actualNum = numbersByRowCol[row][col];
                                
                        if (!actualNum.HasValue || actualNum.Value != expectedNumSlow)
                        {
                            rowOk = false;
                            firstBadIdxSlow = i;
                            break;
                        }
                    }
                            
                    if (!rowOk) break;
                }
        
                // === 检查是否有遗漏的描述行（有图片行的下一行却没有编号）===
                bool hasMissingDescRows = false;
                foreach (int imgRow in imageRows)
                {
                    int descRow = imgRow + 1;
                    if (descRow <= totalRows && !imageRows.Contains(descRow) && !numbersByRow.ContainsKey(descRow))
                    {
                        hasMissingDescRows = true;
                        break;
                    }
                }
        
                if (firstBadIdxSlow >= 0 && !hasMissingDescRows)
                {
                    // 值不连续，且没有遗漏的描述行 → 直接逐个更新编号，跳过结构检查
                    // 多列表格：每个单元格独立递增编号
                    int totalDescRowsToUpdate = sortedRows.Count - firstBadIdxSlow;
                    int totalCellsToUpdate = totalDescRowsToUpdate * colCount;
                    int cellsUpdated = 0;
                            
                    // 计算起始编号值
                    // firstBadIdx 是相对于 sortedRows（scanStart之后）的索引
                    // 所以正确的起始编号应该是 baseNumber + firstBadIdx * colCount + 1
                    int currentNum = baseNumberSlow + firstBadIdxSlow * colCount + 1;
                            
                    // 在更新循环之前切换ScreenUpdating显示进度
                    try { if (app != null) app.ScreenUpdating = true; } catch { }
                    progressCallback?.Invoke("正在更新编号...");
                    System.Windows.Forms.Application.DoEvents();
                    try { if (app != null) app.ScreenUpdating = false; } catch { }
                            
                    // 执行所有更新（不再在循环内切换 ScreenUpdating）
                    for (int i = firstBadIdxSlow; i < sortedRows.Count; i++)
                    {
                        int row = sortedRows[i];
                                
                        // 为每个列设置递增的编号
                        for (int col = 1; col <= colCount; col++)
                        {
                            currentNum++;
                            // 优化3：利用缓存检查，如果单元格当前值已正确，跳过写入
                            if (numbersByRowCol.ContainsKey(row) && 
                                numbersByRowCol[row].ContainsKey(col) && 
                                numbersByRowCol[row][col] == currentNum)
                                continue; // 跳过，节省COM调用
                                    
                            try { UpdateCellNumber(tbl.Cell(row, col), currentNum); }
                            catch { }
                            cellsUpdated++;
                        }
                                
                        // 只做 DoEvents 保持 UI 响应，不切换 ScreenUpdating
                        if ((i - firstBadIdxSlow) % 50 == 0)
                            System.Windows.Forms.Application.DoEvents();
                    }
                            
                    progressCallback?.Invoke(string.Format("编号刷新完成！(更新{0}个编号, 耗时{1:F2}s)",
                        totalCellsToUpdate, (DateTime.Now - startTime).TotalSeconds));
                    return;
                }

                // === 慢速路径：编号值都正确或有遗漏的描述行，检查结构是否需要增删 ===
                // 重要：结构检查必须从表格第1行开始，避免光标位置导致遍漏
                // 临时启用ScreenUpdating让状态栏可见
                try { if (app != null) app.ScreenUpdating = true; } catch { }
                progressCallback?.Invoke("正在检查表格结构...");
                System.Windows.Forms.Application.DoEvents();
                try { if (app != null) app.ScreenUpdating = false; } catch { }

                // 补全 scanStart 前的 InlineShapes
                if (scanStart > 1)
                {
                    try
                    {
                        Range preShapeRange = doc.Range(tbl.Range.Start, tbl.Cell(scanStart, 1).Range.Start);
                        foreach (InlineShape shape in preShapeRange.InlineShapes)
                        {
                            try { imageRows.Add(shape.Range.Cells[1].RowIndex); }
                            catch { }
                        }
                    }
                    catch { }
                }

                // 临时启用ScreenUpdating让状态栏可见
                try { if (app != null) app.ScreenUpdating = true; } catch { }
                progressCallback?.Invoke(string.Format("扫描完成({0}个图片行, {1}个编号行) 已用:{2:F2}s",
                    imageRows.Count, numbersByRow.Count, (DateTime.Now - startTime).TotalSeconds));
                System.Windows.Forms.Application.DoEvents();
                try { if (app != null) app.ScreenUpdating = false; } catch { }

                // === Step 3: 从第1行开始检查结构，修复增删问题 ===
                int addedCount = 0, removedCount = 0;
                bool structureChanged = false;
                int progressInterval = totalRows < 100 ? 20 : (totalRows < 500 ? 50 : 100);
        
                for (int row = 1; row <= totalRows; row++)
                {
                    // 定期更新进度，防止UI卡顿
                    if (row % progressInterval == 0)
                    {
                        // 临时启用ScreenUpdating让状态栏可见
                        try { if (app != null) app.ScreenUpdating = true; } catch { }
                        progressCallback?.Invoke(string.Format("正在检查第 {0}/{1} 行...", row, totalRows));
                        System.Windows.Forms.Application.DoEvents();
                        try { if (app != null) app.ScreenUpdating = false; } catch { }
                    }

                    bool hasImage = imageRows.Contains(row);
                    bool isDescRow = !hasImage && row > 1 && imageRows.Contains(row - 1);
                    // 优化2：用缓存替代 HasNumbering 的500次COM调用
                    // descriptionRows 包含所有图片行下一行的行号
                    // numbersByRowCol 包含所有有编号的描述行的列数据
                    bool hasNum = numbersByRowCol.ContainsKey(row) || 
                                  (row >= scanStart && descriptionRows.Contains(row));
        
                    if (isDescRow && !hasNum)
                    {
                        // 缺少编号 → 插入（编号值会在后续重新编号时确定）
                        for (int col = 1; col <= colCount; col++)
                        {
                            try 
                            { 
                                // 先插入临时编号，后续会重新编号
                                Range r = tbl.Cell(row, col).Range;
                                r.SetRange(r.Start, r.End - 1);
                                r.Text = "0."; // 临时占位
                            }
                            catch { }
                        }
                        addedCount++;
                        structureChanged = true;
                    }
                    else if (!isDescRow && hasNum)
                    {
                        // 多余编号 → 移除
                        // 删除 SEQ 域（向后兼容）
                        for (int col = 1; col <= colCount; col++)
                        {
                            try
                            {
                                Range cellRange = tbl.Cell(row, col).Range;
                                for (int i = cellRange.Fields.Count; i >= 1; i--)
                                {
                                    try 
                                    { 
                                        if (cellRange.Fields[i].Type == WdFieldType.wdFieldSequence)
                                            cellRange.Fields[i].Delete(); 
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                        }
                        // 清理残留文本
                        for (int col = 1; col <= colCount; col++)
                        {
                            try
                            {
                                Range cellRange = tbl.Cell(row, col).Range;
                                cellRange.SetRange(cellRange.Start, cellRange.End - 1);
                                string cellText = CleanCellText(cellRange.Text);
                                if (!string.IsNullOrEmpty(cellText))
                                {
                                    // 去掉编号前缀，保留后面的文本
                                    string cleaned = Regex.Replace(cellText, @"^\d+\.\s*", "");
                                    if (cleaned != cellText) cellRange.Text = cleaned;
                                }
                            }
                            catch { }
                        }
                        numbersByRow.Remove(row);
                        removedCount++;
                        structureChanged = true;
                    }
                }
        
                // === Step 4: 重新编号所有描述行（结构变化或值不连续时） ===
                // 删除行后即使结构没变，值也可能不连续，需要强制更新
                // 多列表格：每个单元格独立递增编号
                if (structureChanged || firstBadIdxSlow >= 0)
                {
                    // 临时启用ScreenUpdating让状态栏可见
                    try { if (app != null) app.ScreenUpdating = true; } catch { }
                    progressCallback?.Invoke("正在重新编号...");
                    System.Windows.Forms.Application.DoEvents();
                    try { if (app != null) app.ScreenUpdating = false; } catch { }
                    
                    try
                    {
                        // 遍历所有行，为描述行重新编号
                        // 关键：scanStart 之前的行保持原样，从 scanStart 开始重新编号
                        // 计算 scanStart 之前应该有多少个编号，用于确定起始编号
                        int numbersBeforeScanStart = 0;
                        for (int row = 1; row < scanStart; row++)
                        {
                            bool hasImage = imageRows.Contains(row);
                            bool isDescRow = !hasImage && row > 1 && imageRows.Contains(row - 1);
                            if (isDescRow)
                                numbersBeforeScanStart += colCount; // 每行有 colCount 个编号
                        }
                        
                        // 从 0 开始计数，scanStart 之前的编号已经计算在 numbersBeforeScanStart 中
                        // 所以当前编号应该是 numbersBeforeScanStart + 1, numbersBeforeScanStart + 2, ...
                        // 但如果 scanStart > 1，我们需要从 baseNumber+1 开始（接续之前的编号）
                        int currentNum;
                        if (scanStart <= 1)
                        {
                            // 全表刷新：从 1 开始
                            currentNum = 0;
                        }
                        else
                        {
                            // 从中间刷新：从 baseNumber 开始（后面会先 ++ 再使用）
                            currentNum = baseNumberSlow;
                        }
                        
                        int cellsProcessed = 0;
                        for (int row = 1; row <= totalRows; row++)
                        {
                            bool hasImage = imageRows.Contains(row);
                            bool isDescRow = !hasImage && row > 1 && imageRows.Contains(row - 1);
                            
                            if (isDescRow)
                            {
                                // scanStart 之前的行：保持原样，只递增计数器
                                // scanStart 及之后的行：重新编号
                                if (row >= scanStart)
                                {
                                    // 为每个列设置递增的编号
                                    for (int col = 1; col <= colCount; col++)
                                    {
                                        currentNum++;
                                        cellsProcessed++;
                                        try { SetCellNumber(tbl.Cell(row, col), currentNum, wdAlignment); }
                                        catch { }
                                    }
                                    
                                    // 每50个单元格做一次DoEvents，不切换ScreenUpdating
                                    if (cellsProcessed % 50 == 0)
                                    {
                                        System.Windows.Forms.Application.DoEvents();
                                    }
                                }
                                else
                                {
                                    // scanStart 之前的行：只递增计数器，不修改单元格
                                    for (int col = 1; col <= colCount; col++)
                                    {
                                        currentNum++;
                                    }
                                }
                            }
                        }
                        
                        // 更新完成后让出UI线程
                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch { }
                }
        
                // === 结果汇报 ===
                if (addedCount == 0 && removedCount == 0)
                {
                    progressCallback?.Invoke(string.Format("编号已是最新，无需更新 (耗时 {0:F2}s)", (DateTime.Now - startTime).TotalSeconds));
                }
                else
                {
                    progressCallback?.Invoke(string.Format("编号刷新完成！(添加{0}, 移除{1}, 耗时{2:F2}s)",
                        addedCount, removedCount, (DateTime.Now - startTime).TotalSeconds));
                }
                
                // 确保UI刷新后再返回
                System.Windows.Forms.Application.DoEvents();
            }
            finally
            {
                try
                {
                    if (app != null) 
                    {
                        // 先恢复ScreenUpdating
                        app.ScreenUpdating = wasScreenUpdating;
                        // 循环DoEvents让Word有充足时间完成屏幕重绘，避免后续UI冻结
                        for (int i = 0; i < 15; i++)
                        {
                            System.Windows.Forms.Application.DoEvents();
                            System.Threading.Thread.Sleep(10);
                        }
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 清除表格自动编号（支持 Word 原生列表编号、SEQ 域和文本编号方式）
        /// </summary>
        public static int ClearTableNumbering(Table tbl, int startRow = 1, Action<string> progressCallback = null)
        {
            if (tbl == null) return 0;

            int originalAlignment = 0;
            Application app = null;
            bool wasScreenUpdating = true;
            Document doc = null;

            try
            {
                try
                {
                    app = tbl.Range.Application;
                    wasScreenUpdating = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                    doc = tbl.Range.Document;
                }
                catch { }

                // 计算清除范围的起始位置
                int clearStartRow = Math.Max(1, startRow);
                int totalRows = tbl.Rows.Count;
                clearStartRow = Math.Min(clearStartRow, totalRows);

                // 优化：如果表格中完全没有 InlineShapes，不需要做任何清理
                // 这处理了空表或全文本表格的情况
                try
                {
                    if (tbl.Range.InlineShapes.Count == 0)
                    {
                        return 0; // 没有图片，不需要清理描述行
                    }
                }
                catch { }

                // 获取 startRow 对应的文档位置，用于判断域是否在清除范围内
                int rangeStartPos = 0;
                try
                {
                    rangeStartPos = tbl.Cell(clearStartRow, 1).Range.Start;
                }
                catch { rangeStartPos = 0; }

                // === 第0步：清除 Word 原生自动编号格式（ListFormat） ===
                // 仅在全表清除时执行（startRow=1），增量模式跳过
                if (clearStartRow <= 1)
                {
                    progressCallback?.Invoke("正在清除列表编号格式...");
                    try
                    {
                        tbl.Range.ListFormat.RemoveNumbers();
                    }
                    catch { }
                }

                // === 第1步优化：只扫描 clearStartRow 之后的域 ===
                progressCallback?.Invoke("正在清除编号...");
                try
                {
                    // 优化3：限制域扫描范围，避免扫描整个表格
                    Range fieldScanRange;
                    if (clearStartRow > 1 && doc != null)
                    {
                        int startPos = tbl.Cell(clearStartRow, 1).Range.Start;
                        fieldScanRange = doc.Range(startPos, tbl.Range.End);
                    }
                    else
                    {
                        fieldScanRange = tbl.Range;
                    }
                    Fields allFields = fieldScanRange.Fields;
                    // 性能优化：如果 Fields 数量为 0，直接跳过遍历
                    if (allFields.Count > 0)
                    {
                        for (int i = allFields.Count; i >= 1; i--)
                        {
                            try
                            {
                                if (allFields[i].Type == WdFieldType.wdFieldSequence)
                                {
                                    // 在删除第一个域前，检测对齐方式
                                    if (originalAlignment == 0)
                                    {
                                        try
                                        {
                                            var paraAlignment = allFields[i].Result.ParagraphFormat.Alignment;
                                            switch (paraAlignment)
                                            {
                                                case WdParagraphAlignment.wdAlignParagraphCenter: originalAlignment = 2; break;
                                                case WdParagraphAlignment.wdAlignParagraphRight: originalAlignment = 3; break;
                                                default: originalAlignment = 1; break;
                                            }
                                        }
                                        catch { originalAlignment = 2; }
                                    }
                                    allFields[i].Delete();
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                // === 第2步优化：跳跃式扫描，只清理描述行（图片行+1）===
                int colCount = tbl.Columns.Count;
                int progressInterval = totalRows < 50 ? 5 : (totalRows < 200 ? 10 : 20);

                // 先找出所有图片行
                var imageRows = new HashSet<int>();
                try
                {
                    Range scanRange = tbl.Range;
                    // 如果 clearStartRow > 1，缩小扫描范围
                    if (clearStartRow > 1 && doc != null)
                    {
                        int startPos = tbl.Cell(clearStartRow, 1).Range.Start;
                        scanRange = doc.Range(startPos, tbl.Range.End);
                    }
                    foreach (InlineShape shape in scanRange.InlineShapes)
                    {
                        try { imageRows.Add(shape.Range.Cells[1].RowIndex); }
                        catch { }
                    }
                }
                catch { }

                // 只清理描述行（图片行+1）
                var rowsToClean = new SortedSet<int>();
                foreach (int imgRow in imageRows)
                {
                    int descRow = imgRow + 1;
                    if (descRow >= clearStartRow && descRow <= totalRows)
                        rowsToClean.Add(descRow);
                }

                // 对标记的行执行清理
                int processedCount = 0;
                foreach (int rowIdx in rowsToClean)
                {
                    try
                    {
                        // 跳过合并单元格行（如文件夹名称标题行）
                        if (tbl.Rows[rowIdx].Cells.Count < colCount) continue;

                        for (int colIdx = 1; colIdx <= colCount; colIdx++)
                        {
                            try
                            {
                                Range cellRange = tbl.Cell(rowIdx, colIdx).Range;

                                // 跳过有图片的单元格
                                if (cellRange.InlineShapes.Count > 0) continue;

                                // 排除 end-of-cell marker 后再读取文本
                                cellRange.SetRange(cellRange.Start, cellRange.End - 1);

                                // 读取单元格纯文本
                                string cellText = (cellRange.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                                if (string.IsNullOrEmpty(cellText)) continue;

                                // 情况1：只剩 "." 或 ". " 或纯点号空格（SEQ域删除后的残留）
                                string stripped = cellText.TrimEnd('.', ' ');
                                if (string.IsNullOrEmpty(stripped) || cellText == "." || cellText == ". ")
                                {
                                    cellRange.Text = "";
                                    continue;
                                }

                                // 情况2：纯文本编号（如 "1."、"2)"、"12"）
                                int? number = ExtractNumberFromCellText(cellText);
                                if (number.HasValue)
                                {
                                    string trimmed = cellText.TrimEnd('.', ')', ' ');
                                    int dummy;
                                    if (int.TryParse(trimmed, out dummy))
                                    {
                                        // 纯编号，检测对齐后清空
                                        if (originalAlignment == 0)
                                        {
                                            try
                                            {
                                                var paraAlignment = cellRange.ParagraphFormat.Alignment;
                                                switch (paraAlignment)
                                                {
                                                    case WdParagraphAlignment.wdAlignParagraphCenter: originalAlignment = 2; break;
                                                    case WdParagraphAlignment.wdAlignParagraphRight: originalAlignment = 3; break;
                                                    default: originalAlignment = 1; break;
                                                }
                                            }
                                            catch { originalAlignment = 2; }
                                        }
                                        cellRange.Text = "";
                                    }
                                    else
                                    {
                                        // 编号+文件名（如 "1. photo1"），只去掉编号前缀
                                        var match = Regex.Match(cellText, @"^\d+[.\)]\s*");
                                        if (match.Success)
                                        {
                                            cellRange.Text = cellText.Substring(match.Length);
                                        }
                                    }
                                }
                            }
                            catch { }
                        }

                        processedCount++;
                        if (progressCallback != null && processedCount % progressInterval == 0)
                        {
                            progressCallback($"正在清除编号... ({processedCount}/{rowsToClean.Count})");
                        }
                    }
                    catch { }
                }
            }
            catch { }
            finally
            {
                try
                {
                    if (app != null) app.ScreenUpdating = wasScreenUpdating;
                }
                catch { }
            }

            return originalAlignment;
        }

        /// <summary>
        /// 在描述行添加自动编号（使用 Word SEQ 域代码方式）
        /// 核心逻辑：前一行有图片 且 当前行无图片 → 当前行是描述行，需要编号
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="doc">文档对象</param>
        /// <param name="startRow">开始行</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="needAutoNumbering">是否需要自动编号</param>
        /// <param name="progressCallback">进度回调</param>
        public static void AddNumberingToDescriptionRows(Table tbl, Document doc,
            int startRow = 1, int alignment = 1, bool needAutoNumbering = false, Action<string> progressCallback = null)
        {
            if (tbl == null || doc == null || !needAutoNumbering) return;

            try
            {
                // 计算起始编号（延续已有编号）
                int startNumber = 1;
                if (startRow > 1)
                {
                    startNumber = CalculateNextSequenceNumber(tbl, startRow);
                }

                // 对齐方式转换
                WdParagraphAlignment wdAlignment;
                switch (alignment)
                {
                    case 2:
                        wdAlignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        break;
                    case 3:
                        wdAlignment = WdParagraphAlignment.wdAlignParagraphRight;
                        break;
                    default:
                        wdAlignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                }

                int colCount = tbl.Columns.Count;
                int totalRows = tbl.Rows.Count;

                // === 计算需要处理的总行数 ===
                int totalToProcess = totalRows - startRow + 1;

                // === 第一遍：预扫描所有行的图片状态 ===
                bool[] rowHasImages = new bool[totalRows + 1]; // 索引0不用

                // 需要从 startRow-1 开始扫描（因为判断描述行需要知道前一行是否有图片）
                int scanStart = Math.Max(1, startRow - 1);
                for (int rowIdx = scanStart; rowIdx <= totalRows; rowIdx++)
                {
                    try
                    {
                        // 定期让出UI线程，防止Word无响应
                        if ((rowIdx - scanStart) % 100 == 0 && rowIdx > scanStart)
                        {
                            System.Windows.Forms.Application.DoEvents();
                        }

                        // 跳过合并行
                        if (tbl.Rows[rowIdx].Cells.Count < colCount) continue;

                        // 用行级 Range.Text 检查内联图片标记字符（\x01），比 InlineShapes.Count 快得多
                        // Word 中每个内联图片在 Range.Text 中表示为 \x01 字符
                        string rowText = tbl.Rows[rowIdx].Range.Text;
                        if (rowText != null && rowText.IndexOf('\x01') >= 0)
                        {
                            rowHasImages[rowIdx] = true;
                        }
                    }
                    catch { }
                }

                // === 第二遍：编号（使用纯文本编号，传入递增的编号值）===
                int currentNumber = startNumber;
                for (int rowIdx = startRow; rowIdx <= totalRows; rowIdx++)
                {
                    try
                    {
                        // 定期让出UI线程并更新进度
                        if ((rowIdx - startRow) % 50 == 0 && rowIdx > startRow)
                        {
                            if (progressCallback != null && totalToProcess > 0)
                            {
                                int percent = (rowIdx - startRow) * 100 / totalToProcess;
                                progressCallback($"正在编号... {percent}%");
                            }
                            System.Windows.Forms.Application.DoEvents();
                        }

                        // 跳过合并行
                        if (tbl.Rows[rowIdx].Cells.Count < colCount) continue;

                        bool currentRowHasImages = rowHasImages[rowIdx];
                        bool prevHasImages = rowIdx > 1 ? rowHasImages[rowIdx - 1] : false;

                        // 描述行条件：当前行无图片，前一行有图片
                        if (!currentRowHasImages && prevHasImages)
                        {
                            // 为描述行的每个单元格插入纯文本编号，每列独立递增
                            // 例如2列表格：描述行1 col1=1, col2=2; 描述行2 col1=3, col2=4
                            for (int colIdx = 1; colIdx <= colCount; colIdx++)
                            {
                                try
                                {
                                    InsertNumberText(tbl, rowIdx, colIdx, wdAlignment, currentNumber);
                                }
                                catch { }
                                currentNumber++; // 每列递增
                            }
                        }
                    }
                    catch { }
                }

                // 纯文本编号无需更新域，直接完成
                progressCallback?.Invoke("编号完成 100%");
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 计算下一个 SEQ 编号起始值
        /// </summary>
        public static int CalculateNextSequenceNumber(Table tbl, int startRow, int tableColCount = 2)
        {
            // 优化：如果 startRow <= 1，直接从 1 开始编号，无需搜索
            if (startRow <= 1)
            {
                return 1;
            }

            // 极简化实现：只搜索最近2行，只用纯文本匹配
            // 避免调用 ExtractNumberFromCell（它有 Fields 遍历开销）
            for (int row = startRow - 1; row >= Math.Max(1, startRow - 2); row--)
            {
                try
                {
                    Cell cell = tbl.Cell(row, tableColCount);
                    Range r = cell.Range;
                    r.SetRange(r.Start, r.End - 1);
                    string text = (r.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                    var m = Regex.Match(text, @"^(\d+)\.");
                    if (m.Success) return int.Parse(m.Groups[1].Value) + 1;
                }
                catch { }
            }

            // 如果最近2行没找到，返回1（从1开始编号）
            return 1;
        }

        /// <summary>
        /// 从单元格提取编号值（支持纯文本和旧版SEQ域两种格式，向后兼容）
        /// </summary>
        private static int? ExtractNumberFromCell(Cell cell)
        {
            try
            {
                // 优化：先检查纯文本编号（快速路径，G方案文档）
                // 纯文本检查比 Fields 遍历快得多
                Range r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                string text = (r.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                var m = Regex.Match(text, @"^(\d+)\.");
                if (m.Success && int.TryParse(m.Groups[1].Value, out int val))
                    return val;
                
                // 再检查 SEQ 域（慢速路径，旧文档兼容）
                // 注意：在 G 方案下大部分文档不会走到这里
                foreach (Field f in cell.Range.Fields)
                {
                    if (f.Type == WdFieldType.wdFieldSequence)
                    {
                        string resultText = f.Result != null ? f.Result.Text : null;
                        if (!string.IsNullOrEmpty(resultText))
                        {
                            var match = Regex.Match(resultText.Trim(), @"^(\d+)");
                            if (match.Success && int.TryParse(match.Groups[1].Value, out int num))
                                return num;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 检查单元格是否有编号（SEQ域或纯文本格式）
        /// </summary>
        private static bool HasNumbering(Cell cell)
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
            catch { }
            return false;
        }

        /// <summary>
        /// 轻量级编号更新：仅替换文本中的数字前缀
        /// 适用于已确认是纯文本编号格式的单元格（快速路径使用）
        /// </summary>
        private static void UpdateCellNumber(Cell cell, int number)
        {
            try
            {
                Range r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                string text = (r.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                // 只替换开头的数字
                string newText = Regex.Replace(text, @"^\d+", number.ToString());
                if (newText != text)
                {
                    r.Text = newText;
                }
            }
            catch { }
        }

        /// <summary>
        /// 将单元格编号设为指定值（纯文本方式）
        /// 如果有旧版SEQ域则先删除，然后写入 "N. 原文" 格式
        /// </summary>
        private static void SetCellNumber(Cell cell, int number, WdParagraphAlignment alignment)
        {
            try
            {
                Range r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                
                // 删除所有 SEQ 域（向后兼容）
                for (int i = r.Fields.Count; i >= 1; i--)
                {
                    try
                    {
                        if (r.Fields[i].Type == WdFieldType.wdFieldSequence)
                            r.Fields[i].Delete();
                    }
                    catch { }
                }
                
                // 重新获取 Range（删除域后原 Range 失效）
                r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                
                // 提取编号之后的文本部分
                string text = r.Text ?? "";
                text = text.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                string suffix = Regex.Replace(text, @"^\d*\.?\s*", ""); // 去掉旧编号前缀
                
                // 写入新编号
                r.Text = string.IsNullOrEmpty(suffix) ? number + "." : number + ". " + suffix;
                
                // 重新获取 Range 设置对齐（Text赋值后Range失效）
                r = cell.Range;
                r.SetRange(r.Start, r.End - 1);
                r.ParagraphFormat.Alignment = alignment;
            }
            catch { }
        }

        /// <summary>
        /// 从单元格中提取 SEQ 域的编号值（向后兼容，内部调用 ExtractNumberFromCell）
        /// </summary>
        private static int? ExtractSeqNumberFromCell(Cell cell)
        {
            return ExtractNumberFromCell(cell);
        }

        /// <summary>
        /// 在指定单元格插入纯文本编号（替代旧版 SEQ 域方式）
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="number">编号值</param>
        /// <param name="skipFieldCheck">是否跳过 SEQ 域检查（新插入的行可设为 true 以提升性能）</param>
        /// <param name="numberPosition">编号位置（1=在前，2=在后）</param>
        public static void InsertNumberText(Table tbl, int row, int col,
            WdParagraphAlignment alignment, int number, bool skipFieldCheck = false, int numberPosition = 1)
        {
            try
            {
                Cell cell = tbl.Cell(row, col);
                Range cellRange = cell.Range;
                cellRange.SetRange(cellRange.Start, cellRange.End - 1);

                // 获取现有文本内容（如果有）
                string existingText = cellRange.Text ?? "";
                existingText = existingText.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                // 去掉可能的旧编号前缀（格式：数字+点号+空格，如 "1. "、"12. "）
                // 注意：只匹配完整的编号格式，避免误删描述文本开头的数字
                existingText = Regex.Replace(existingText, @"^\d+\.\s+", "");

                // 删除所有旧 SEQ 域（向后兼容，但新插入的行可跳过以提升性能）
                if (!skipFieldCheck)
                {
                    for (int i = cellRange.Fields.Count; i >= 1; i--)
                    {
                        try
                        {
                            if (cellRange.Fields[i].Type == WdFieldType.wdFieldSequence)
                                cellRange.Fields[i].Delete();
                        }
                        catch { }
                    }

                    // 重新获取 Range（删除域后原 Range 可能失效）
                    cellRange = cell.Range;
                    cellRange.SetRange(cellRange.Start, cellRange.End - 1);
                }

                // 根据编号位置生成新文本
                string newText;
                if (numberPosition == 2 && !string.IsNullOrEmpty(existingText))
                {
                    // 编号在描述后面：描述-编号
                    newText = existingText + "-" + number;
                }
                else
                {
                    // 编号在描述前面（默认）：编号. 描述
                    if (!string.IsNullOrEmpty(existingText))
                        newText = number + ". " + existingText;
                    else
                        newText = number + ".";
                }

                cellRange.Text = newText;

                // 重新获取 Range 设置对齐（Text赋值后Range失效）
                cellRange = cell.Range;
                cellRange.SetRange(cellRange.Start, cellRange.End - 1);
                cellRange.ParagraphFormat.Alignment = alignment;
            }
            catch { }
        }

        /// <summary>
        /// 兼容包装：旧签名的 InsertSeqField，内部转为纯文本编号
        /// 注意：调用方需要自行维护编号计数（startNumber 参数传入当前应显示的编号值）
        /// </summary>
        public static void InsertSeqField(Table tbl, int rowIdx, int colIdx,
            WdParagraphAlignment alignment, ref bool isFirstSeqField, int startNumber)
        {
            // startNumber 在旧接口中就是当前应显示的编号值
            InsertNumberText(tbl, rowIdx, colIdx, alignment, startNumber);
            isFirstSeqField = false;
        }

        /// <summary>
        /// 从单元格文本中提取编号值（如 "1." 提取出 1）
        /// </summary>
        private static int? ExtractNumberFromCellText(string cellText)
        {
            if (string.IsNullOrEmpty(cellText)) return null;

            // 匹配以数字开头，后跟点号或括号的格式（如 "1.", "2)", "3 " 等）
            var match = Regex.Match(cellText, @"^(\d+)");
            if (match.Success)
            {
                if (int.TryParse(match.Groups[1].Value, out int number))
                {
                    return number;
                }
            }
            return null;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 清理单元格文本
        /// </summary>
        private static string CleanCellText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("\r", "").Replace("\n", "")
                .Replace("\t", "").Replace("\a", "").Replace("\u00A0", "").Trim();
        }

        #endregion
    }
}
