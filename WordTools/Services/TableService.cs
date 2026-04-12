using System;
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
            if (targetCell == null) return false;

            try
            {
                // 检查是否已有图片
                if (targetCell.Range.InlineShapes.Count > 0)
                {
                    return false;
                }

                // 检查单元格是否包含 SEQ 域（编号域）
                bool hasSeqField = false;
                foreach (Field field in targetCell.Range.Fields)
                {
                    if (field.Type == WdFieldType.wdFieldSequence)
                    {
                        hasSeqField = true;
                        break;
                    }
                }

                // 如果包含 SEQ 域，清除域使其可复用
                if (hasSeqField)
                {
                    try
                    {
                        // 删除所有 SEQ 域
                        int fieldCount = targetCell.Range.Fields.Count;
                        for (int i = fieldCount; i >= 1; i--)
                        {
                            if (targetCell.Range.Fields[i].Type == WdFieldType.wdFieldSequence)
                            {
                                targetCell.Range.Fields[i].Delete();
                            }
                        }
                        // 清除剩余文本
                        targetCell.Range.Text = "";
                    }
                    catch
                    {
                        // 清除失败则不适合插入
                        return false;
                    }
                    return true;
                }

                // 获取单元格文本并清理
                string cellText = CleanCellText(targetCell.Range.Text);

                // 空单元格适合插入
                if (string.IsNullOrEmpty(cellText))
                {
                    return true;
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
                        return false;
                    }
                    return true;
                }

                return true;
            }
            catch
            {
                return true;
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

                    // 检查第1列
                    bool col1Suitable = IsCellSuitableForImage(tbl.Cell(row, 1));
                    
                    // 如果第1列不适合，跳过整行
                    if (!col1Suitable) continue;

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

        #endregion

        #region 表格操作

        /// <summary>
        /// 确保表格行存在
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
        public static void InsertFileNameDescriptionRow(Table tbl, ref int rowIndex, string[] fileNames)
        {
            if (tbl == null) return;

            try
            {
                EnsureRowExists(tbl, rowIndex);
                AdjustTableColumns(tbl, 2);

                // 插入文件名到对应列
                for (int i = 0; i < Math.Min(fileNames.Length, 2); i++)
                {
                    string baseName = FileService.GetFileNameWithoutExtension(fileNames[i]);
                    var cell = tbl.Cell(rowIndex, i + 1);
                    cell.Range.Text = baseName;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                // 如果只有一个文件，第二列留空
                if (fileNames.Length < 2)
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
        /// 刷新整个表格的编号（清除并重新添加）
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="doc">文档对象</param>
        /// <param name="alignment">对齐方式（1=居左, 2=居中, 3=居右）</param>
        /// <param name="progressCallback">进度回调</param>
        public static void RefreshTableNumbering(Table tbl, Document doc, int alignment = 2, Action<string> progressCallback = null)
        {
            if (tbl == null || doc == null) return;

            // 1. 清除编号（带进度）
            progressCallback?.Invoke("正在清除编号...");
            ClearTableNumbering(tbl, 1, progressCallback);

            // 2. 添加编号（带进度）
            progressCallback?.Invoke("正在添加编号...");
            AddNumberingToDescriptionRows(tbl, doc, 1, alignment, true, progressCallback);

            progressCallback?.Invoke("编号刷新完成！");
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

            try
            {
                try
                {
                    app = tbl.Range.Application;
                    wasScreenUpdating = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                }
                catch { }

                // 计算清除范围的起始位置
                int clearStartRow = Math.Max(1, startRow);
                int totalRows = tbl.Rows.Count;
                clearStartRow = Math.Min(clearStartRow, totalRows);

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

                // === 第1步：删除 startRow 之后的 SEQ 域 ===
                progressCallback?.Invoke("正在清除 SEQ 域...");
                try
                {
                    Fields allFields = tbl.Range.Fields;
                    for (int i = allFields.Count; i >= 1; i--)
                    {
                        try
                        {
                            if (allFields[i].Type == WdFieldType.wdFieldSequence)
                            {
                                // 只删除位于 startRow 之后的域
                                if (allFields[i].Result.Start >= rangeStartPos)
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
                        }
                        catch { }
                    }
                }
                catch { }

                // === 第2步：逐行清理残留文本（域删除后的 "."、纯文本编号等） ===
                int colCount = tbl.Columns.Count;
                int progressInterval = totalRows < 50 ? 5 : (totalRows < 200 ? 10 : 20);

                for (int rowIdx = clearStartRow; rowIdx <= totalRows; rowIdx++)
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

                                // 读取单元格纯文本
                                string cellText = CleanCellText(cellRange.Text);
                                if (string.IsNullOrEmpty(cellText)) continue;

                                // 排除 end-of-cell marker 后再操作文本
                                cellRange.SetRange(cellRange.Start, cellRange.End - 1);

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

                        if (progressCallback != null && rowIdx % progressInterval == 0)
                        {
                            progressCallback($"正在清除编号... ({rowIdx}/{totalRows})");
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

                bool isFirstSeqField = true;
                int colCount = tbl.Columns.Count;
                int totalRows = tbl.Rows.Count;

                // 计算进度更新间隔
                int progressInterval = totalRows < 50 ? 5 : (totalRows < 200 ? 10 : 20);

                // 单遍扫描 + 编号
                bool prevRowHasImages = false;

                for (int rowIdx = startRow; rowIdx <= totalRows; rowIdx++)
                {
                    try
                    {
                        // 跳过合并单元格行（如文件夹名称标题行）
                        bool isMergedRow = false;
                        try
                        {
                            if (tbl.Rows[rowIdx].Cells.Count < colCount)
                            {
                                isMergedRow = true;
                            }
                        }
                        catch { isMergedRow = true; }

                        if (isMergedRow)
                        {
                            prevRowHasImages = false;
                            continue;
                        }

                        // 检查当前行是否有图片（找到第一个就 break）
                        bool currentRowHasImages = false;
                        for (int c = 1; c <= colCount; c++)
                        {
                            try
                            {
                                if (tbl.Cell(rowIdx, c).Range.InlineShapes.Count > 0)
                                {
                                    currentRowHasImages = true;
                                    break;  // 找到一个就够了
                                }
                            }
                            catch { }
                        }

                        // 判断是否为描述行：前一行有图片 且 当前行无图片
                        if (!currentRowHasImages && prevRowHasImages)
                        {
                            // 为描述行的每个单元格插入 SEQ 域
                            for (int colIdx = 1; colIdx <= colCount; colIdx++)
                            {
                                try
                                {
                                    InsertSeqField(tbl, rowIdx, colIdx, wdAlignment, ref isFirstSeqField, startNumber);
                                }
                                catch { }
                            }
                        }

                        prevRowHasImages = currentRowHasImages;

                        // 进度回调
                        if (progressCallback != null && (rowIdx - startRow) % progressInterval == 0)
                        {
                            progressCallback($"正在添加编号... ({rowIdx}/{totalRows})");
                        }
                    }
                    catch
                    {
                        prevRowHasImages = false;
                    }
                }

                // 更新域：增量模式只更新 startRow 之后的范围
                try
                {
                    if (startRow > 1)
                    {
                        // 增量模式：缩小更新范围，只更新 startRow 之后的域
                        Range updateRange = tbl.Cell(startRow, 1).Range;
                        updateRange.SetRange(updateRange.Start, tbl.Range.End);
                        updateRange.Fields.Update();
                    }
                    else
                    {
                        // 全表模式
                        tbl.Range.Fields.Update();
                    }
                }
                catch { }
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 计算下一个 SEQ 编号起始值
        /// </summary>
        private static int CalculateNextSequenceNumber(Table tbl, int startRow)
        {
            int maxNumber = 0;
            int tableColCount = tbl.Columns.Count;

            for (int checkRow = startRow - 1; checkRow >= 1; checkRow--)
            {
                try
                {
                    int checkColCount = tbl.Rows[checkRow].Cells.Count;
                    // 跳过合并单元格行（如文件夹标题行），避免误读标题文本中的数字
                    if (checkColCount < tableColCount) continue;

                    // 从右往左查找最后一个有编号的单元格
                    for (int col = checkColCount; col >= 1; col--)
                    {
                        try
                        {
                            var cell = tbl.Cell(checkRow, col);

                            // 优先检查 SEQ 域
                            int? seqNumber = ExtractSeqNumberFromCell(cell);
                            if (seqNumber.HasValue && seqNumber.Value > maxNumber)
                            {
                                maxNumber = seqNumber.Value;
                            }

                            // 兼容检查纯文本编号（仅匹配纯编号格式如 "1."、"2)"、"3"）
                            if (maxNumber == 0)
                            {
                                string cellText = CleanCellText(cell.Range.Text);
                                if (!string.IsNullOrEmpty(cellText))
                                {
                                    // 只匹配纯数字或 "数字." / "数字)" 格式，排除 "336-glass bottles" 等文件夹标题
                                    string trimmed = cellText.TrimEnd('.', ')', ' ');
                                    int parsedNumber;
                                    if (int.TryParse(trimmed, out parsedNumber) && parsedNumber > maxNumber)
                                    {
                                        maxNumber = parsedNumber;
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    if (maxNumber > 0) break;
                }
                catch { }
            }

            return maxNumber > 0 ? maxNumber + 1 : 1;
        }

        /// <summary>
        /// 从单元格中提取 SEQ 域的编号值
        /// </summary>
        private static int? ExtractSeqNumberFromCell(Cell cell)
        {
            try
            {
                foreach (Field field in cell.Range.Fields)
                {
                    if (field.Type == WdFieldType.wdFieldSequence)
                    {
                        string resultText = field.Result?.Text;
                        if (!string.IsNullOrEmpty(resultText))
                        {
                            // 提取数字部分
                            var match = Regex.Match(resultText.Trim(), @"^(\d+)");
                            if (match.Success && int.TryParse(match.Groups[1].Value, out int number))
                            {
                                return number;
                            }
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 在指定单元格插入 SEQ 域
        /// </summary>
        private static void InsertSeqField(Table tbl, int rowIdx, int colIdx,
            WdParagraphAlignment alignment, ref bool isFirstSeqField, int startNumber)
        {
            try
            {
                // 1. 获取单元格 Range 和起止位置（只获取一次 Cell，后续复用位置信息）
                Range cellRange = tbl.Cell(rowIdx, colIdx).Range;
                int cellStart = cellRange.Start;
                int cellEnd = cellRange.End - 1; // 排除 end-of-cell marker
        
                // 2. 读取现有文本（复用 cellRange）
                cellRange.SetRange(cellStart, cellEnd);
                string existingText = CleanCellText(cellRange.Text);
        
                // 3. 清理原内容：去掉编号前缀和残留的点号，只保留文件名
                if (!string.IsNullOrEmpty(existingText))
                {
                    // 先去掉编号前缀（如 "1. ", "2) "）
                    existingText = Regex.Replace(existingText, @"^\d+[.)\]\s*", "");
                    // 再去掉残留的点号和空格（如刷新编号后留下的 ". " 或 "."）
                    existingText = existingText.TrimStart('.', ' ');
                }
        
                // 4. 清空单元格内容（复用 cellRange）
                cellRange.SetRange(cellStart, cellEnd);
                cellRange.Text = "";
        
                // 5. 构建 SEQ 域参数
                string fieldText;
                if (isFirstSeqField && startNumber > 1)
                {
                    fieldText = $"PhotoNum \\r {startNumber}";
                }
                else
                {
                    fieldText = "PhotoNum";
                }
        
                // 6. 在单元格起始位置插入 SEQ 域
                // 插入域后 Range 会失效，需要重新定位
                cellRange.SetRange(cellStart, cellStart);
                cellRange.Fields.Add(cellRange, WdFieldType.wdFieldSequence, fieldText, false);
                isFirstSeqField = false;
        
                // 7. 在域后追加 "." 和文件名（域插入后必须重新获取 Range）
                Range afterField = tbl.Cell(rowIdx, colIdx).Range;
                afterField.SetRange(afterField.End - 1, afterField.End - 1);
        
                if (!string.IsNullOrEmpty(existingText))
                {
                    afterField.Text = ". " + existingText;
                }
                else
                {
                    afterField.Text = ".";
                }
        
                // 8. 设置单元格格式（需要重新获取 Range）
                Range fullRange = tbl.Cell(rowIdx, colIdx).Range;
                fullRange.SetRange(fullRange.Start, fullRange.End - 1);
                fullRange.ParagraphFormat.Alignment = alignment;
                tbl.Cell(rowIdx, colIdx).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
            catch { }
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
