using System;
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

                // 检查是否使用了自动编号，如果是则清除编号使其可复用
                if (targetCell.Range.Paragraphs.Count > 0)
                {
                    var para = targetCell.Range.Paragraphs[1];
                    if (para.Range.ListFormat.ListType != WdListType.wdListNoNumbering)
                    {
                        try
                        {
                            targetCell.Range.ListFormat.RemoveNumbers();
                            targetCell.Range.Text = "";
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }

                // 获取单元格文本并清理
                string cellText = targetCell.Range.Text ?? "";
                cellText = cellText.Replace("\r", "").Replace("\n", "")
                    .Replace("\t", "").Replace("\a", "").Replace("\u00A0", "").Trim();

                // 空单元格适合插入
                if (string.IsNullOrEmpty(cellText))
                {
                    return true;
                }

                // 检查是否为序号格式（纯数字或数字+标点）
                // 如果是序号，清除文本使其可复用
                if (char.IsDigit(cellText[0]))
                {
                    // 检查是否只包含数字和常见序号标点
                    bool isValidNumber = true;
                    foreach (char c in cellText)
                    {
                        if (!char.IsDigit(c) && c != '.' && c != ')' && c != '(' && c != '-' && c != ' ')
                        {
                            isValidNumber = false;
                            break;
                        }
                    }
                    if (isValidNumber)
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
                    if (preferredCol == 2 && col2Suitable)
                    {
                        foundRow = row;
                        foundCol = 2;
                        return true;
                    }
                    else if (preferredCol == 2 && col1Suitable)
                    {
                        foundRow = row;
                        foundCol = 1;
                        return true;
                    }
                    else if (preferredCol == 1 && col1Suitable)
                    {
                        foundRow = row;
                        foundCol = 1;
                        return true;
                    }
                    else if (preferredCol == 1 && col2Suitable)
                    {
                        foundRow = row;
                        foundCol = 2;
                        return true;
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
                    cell.Range.ListFormat.RemoveNumbers();
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
        /// 清除表格自动编号
        /// </summary>
        /// <param name="tbl">表格对象</param>
        /// <param name="startRow">开始行</param>
        /// <returns>原来的编号对齐方式（1=居左, 2=居中, 3=居右）</returns>
        public static int ClearTableNumbering(Table tbl, int startRow = 1)
        {
            if (tbl == null) return 0;

            int originalAlignment = 0;

            try
            {
                int clearStartRow = startRow == 1 ? 1 : startRow;

                // 检测原来的编号对齐方式
                if (tbl.Rows.Count > 0)
                {
                    int checkRow = clearStartRow > 1 && clearStartRow <= tbl.Rows.Count 
                        ? clearStartRow - 1 : 1;

                    try
                    {
                        if (tbl.Cell(checkRow, 1).Range.ListFormat.ListType != WdListType.wdListNoNumbering)
                        {
                            var alignment = tbl.Cell(checkRow, 1).Range.ParagraphFormat.Alignment;
                            switch (alignment)
                            {
                                case WdParagraphAlignment.wdAlignParagraphLeft:
                                    originalAlignment = 1;
                                    break;
                                case WdParagraphAlignment.wdAlignParagraphCenter:
                                    originalAlignment = 2;
                                    break;
                                case WdParagraphAlignment.wdAlignParagraphRight:
                                    originalAlignment = 3;
                                    break;
                                default:
                                    originalAlignment = 1;
                                    break;
                            }
                        }
                    }
                    catch
                    {
                        // 忽略错误
                    }
                }

                // 清除编号
                for (int rowIdx = clearStartRow; rowIdx <= tbl.Rows.Count; rowIdx++)
                {
                    try
                    {
                        var row = tbl.Rows[rowIdx];
                        bool hadNumbering = row.Range.ListFormat.ListType != WdListType.wdListNoNumbering;
                        row.Range.ListFormat.RemoveNumbers();
                        
                        // 清除编号行的单元格内容
                        try
                        {
                            int colCount = row.Cells.Count;
                            for (int colIdx = 1; colIdx <= colCount; colIdx++)
                            {
                                var cell = tbl.Cell(rowIdx, colIdx);
                                // 只清除没有图片的单元格
                                if (cell.Range.InlineShapes.Count > 0)
                                {
                                    continue;
                                }
                                
                                // 如果该行原来有列表格式编号，直接清除
                                if (hadNumbering)
                                {
                                    cell.Range.Text = "";
                                    continue;
                                }
                                
                                // 检查是否为纯文本形式的序号（如 "1.", "2)", "3" 等）
                                string cellText = cell.Range.Text ?? "";
                                cellText = cellText.Replace("\r", "").Replace("\n", "")
                                    .Replace("\t", "").Replace("\a", "").Replace("\u00A0", "").Trim();
                                
                                if (!string.IsNullOrEmpty(cellText) && char.IsDigit(cellText[0]))
                                {
                                    bool isNumberOnly = true;
                                    foreach (char c in cellText)
                                    {
                                        if (!char.IsDigit(c) && c != '.' && c != ')' && c != '(' && c != '-' && c != ' ')
                                        {
                                            isNumberOnly = false;
                                            break;
                                        }
                                    }
                                    if (isNumberOnly)
                                    {
                                        cell.Range.Text = "";
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // 忽略单元格清理错误
                        }
                    }
                    catch
                    {
                        // 忽略错误
                    }
                }
            }
            catch
            {
                // 忽略错误
            }

            return originalAlignment;
        }

        /// <summary>
        /// 在描述行添加自动编号
        /// </summary>
        public static void AddNumberingToDescriptionRows(Table tbl, Document doc, 
            int startRow = 1, int alignment = 1, bool needAutoNumbering = false)
        {
            if (tbl == null || doc == null || !needAutoNumbering) return;

            try
            {
                // 找到包含图片的最后一行
                int lastImageRow = 0;
                int consecutiveEmptyRows = 0;

                for (int rowIdx = startRow; rowIdx <= tbl.Rows.Count; rowIdx++)
                {
                    try
                    {
                        int colCount = tbl.Rows[rowIdx].Cells.Count;
                        if (colCount >= 2)
                        {
                            bool hasImage = tbl.Cell(rowIdx, 1).Range.InlineShapes.Count > 0 ||
                                          tbl.Cell(rowIdx, 2).Range.InlineShapes.Count > 0;

                            if (hasImage)
                            {
                                lastImageRow = rowIdx;
                                consecutiveEmptyRows = 0;
                            }
                            else
                            {
                                consecutiveEmptyRows++;
                                if (consecutiveEmptyRows >= 2 && lastImageRow > 0)
                                {
                                    break;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 忽略错误
                    }
                }

                if (lastImageRow == 0) return;

                int endRow = Math.Min(lastImageRow + 1, tbl.Rows.Count);

                // 计算编号起始值：查找 startRow 之前已有的编号，延续编号
                int numberStartAt = 1;
                if (startRow > 1)
                {
                    for (int checkRow = startRow - 1; checkRow >= 1; checkRow--)
                    {
                        try
                        {
                            int checkColCount = tbl.Rows[checkRow].Cells.Count;
                            if (checkColCount < 2) continue;

                            // 从右往左查找最后一个有编号的单元格，获取其编号值
                            for (int col = checkColCount; col >= 1; col--)
                            {
                                try
                                {
                                    var cellRange = tbl.Cell(checkRow, col).Range;
                                    if (cellRange.ListFormat.ListType != WdListType.wdListNoNumbering)
                                    {
                                        int listValue = cellRange.ListFormat.ListValue;
                                        if (listValue > 0)
                                        {
                                            numberStartAt = listValue + 1;
                                        }
                                        break;
                                    }
                                }
                                catch { }
                            }

                            if (numberStartAt > 1) break;
                        }
                        catch { }
                    }
                }

                // 创建自定义列表模板
                ListTemplate sharedListTemplate = null;
                bool isFirstDescriptionRow = true;

                // 遍历并添加编号
                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++)
                {
                    try
                    {
                        int colCount = tbl.Rows[rowIdx].Cells.Count;
                        if (colCount < 2) continue;

                        // 检查是否是描述行
                        bool cell1HasImage = tbl.Cell(rowIdx, 1).Range.InlineShapes.Count > 0;
                        bool cell2HasImage = tbl.Cell(rowIdx, 2).Range.InlineShapes.Count > 0;

                        if (!cell1HasImage && !cell2HasImage)
                        {
                            // 检查前一行是否有图片
                            bool prevRowHasImage = false;
                            if (rowIdx > startRow)
                            {
                                try
                                {
                                    prevRowHasImage = tbl.Cell(rowIdx - 1, 1).Range.InlineShapes.Count > 0 ||
                                                     tbl.Cell(rowIdx - 1, 2).Range.InlineShapes.Count > 0;
                                }
                                catch
                                {
                                    // 忽略错误
                                }
                            }

                            if (prevRowHasImage || rowIdx == startRow + 1)
                            {
                                // 这是描述行，添加编号
                                var rowRange = tbl.Rows[rowIdx].Range;

                                if (isFirstDescriptionRow)
                                {
                                    // 创建新列表模板，使用计算的起始编号值
                                    var customListTemplate = doc.ListTemplates.Add(true);
                                    var level = customListTemplate.ListLevels[1];
                                    level.NumberFormat = "%1.";
                                    level.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                                    level.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
                                    level.StartAt = numberStartAt;

                                    rowRange.ListFormat.ApplyListTemplate(
                                        customListTemplate, 
                                        ContinuePreviousList: false, 
                                        ApplyTo: WdListApplyTo.wdListApplyToWholeList);
                                    sharedListTemplate = customListTemplate;
                                    isFirstDescriptionRow = false;
                                }
                                else if (sharedListTemplate != null)
                                {
                                    rowRange.ListFormat.ApplyListTemplate(
                                        sharedListTemplate,
                                        ContinuePreviousList: true,
                                        ApplyTo: WdListApplyTo.wdListApplyToWholeList);
                                }

                                // 设置对齐方式
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
                                rowRange.ParagraphFormat.Alignment = wdAlignment;
                            }
                        }
                    }
                    catch
                    {
                        // 忽略单行错误
                    }
                }
            }
            catch
            {
                // 忽略错误
            }
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
