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
        
                // 扫描从光标前一行开始（需判断“图片行→描述行”配对）
                int scanStart = Math.Max(1, cursorRow - 1);
        
                progressCallback?.Invoke(string.Format("从第 {0} 行开始检查编号...", scanStart));
                System.Windows.Forms.Application.DoEvents();
        
                // === Step 2: 扫描 SEQ 域（快速，只遍历 Fields） ===
                var seqFieldsByRow = new Dictionary<int, List<Field>>();
                Range scanRange = null;
                int fieldCount = 0;

                try
                {
                    scanRange = doc.Range(tbl.Cell(scanStart, 1).Range.Start, tbl.Range.End);

                    foreach (Field f in scanRange.Fields)
                    {
                        try
                        {
                            if (f.Type == WdFieldType.wdFieldSequence)
                            {
                                int rowIdx = f.Result.Cells[1].RowIndex;
                                if (!seqFieldsByRow.ContainsKey(rowIdx))
                                    seqFieldsByRow[rowIdx] = new List<Field>();
                                seqFieldsByRow[rowIdx].Add(f);
                            }
                        }
                        catch { }
                        // 每100个域让出UI线程
                        fieldCount++;
                        if (fieldCount % 100 == 0)
                            System.Windows.Forms.Application.DoEvents();
                    }
                }
                catch { }

                // === 快速路径：检查域值连续性，不连续则直接批量更新 ===
                // 最常见场景：删行后值不对，无需扫描 InlineShapes
                var sortedRows = new List<int>(seqFieldsByRow.Keys);
                sortedRows.Sort();

                int firstBadIdx = -1;
                int prevVal = -1;
                for (int i = 0; i < sortedRows.Count; i++)
                {
                    try
                    {
                        int num = int.Parse(seqFieldsByRow[sortedRows[i]][0].Result.Text.Trim());
                        if (prevVal >= 0 && num != prevVal + 1)
                        {
                            firstBadIdx = i;
                            break;
                        }
                        prevVal = num;
                    }
                    catch { firstBadIdx = i; break; }
                }

                // === 扫描全表 InlineShapes（为后续结构检查做准备）===
                var imageRows = new HashSet<int>();
                int shapeCount = 0;
                try
                {
                    foreach (InlineShape shape in tbl.Range.InlineShapes)
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

                // === 检查是否有遗漏的描述行（有图片行的下一行却没有SEQ域）===
                bool hasMissingDescRows = false;
                foreach (int imgRow in imageRows)
                {
                    int descRow = imgRow + 1;
                    if (descRow <= totalRows && !imageRows.Contains(descRow) && !seqFieldsByRow.ContainsKey(descRow))
                    {
                        hasMissingDescRows = true;
                        break;
                    }
                }

                if (firstBadIdx >= 0 && !hasMissingDescRows)
                {
                    // 值不连续，且没有遗漏的描述行 → 直接批量更新，跳过结构检查
                    int totalToUpdate = sortedRows.Count - firstBadIdx;
                    progressCallback?.Invoke(string.Format("正在批量更新 {0} 个编号值...", totalToUpdate));
                    System.Windows.Forms.Application.DoEvents();

                    try
                    {
                        int firstBadRow = sortedRows[firstBadIdx];
                        Range updateRange = doc.Range(
                            tbl.Cell(firstBadRow, 1).Range.Start,
                            tbl.Range.End);
                        updateRange.Fields.Update();
                    }
                    catch { }

                    progressCallback?.Invoke(string.Format("编号刷新完成！(更新{0}个域, 耗时{1:F2}s)",
                        totalToUpdate, (DateTime.Now - startTime).TotalSeconds));
                    return;
                }

                // === 慢速路径：域值都正确或有遗漏的描述行，检查结构是否需要增删 ===
                // 重要：结构检查必须从表格第1行开始，避免光标位置导致遍漏
                progressCallback?.Invoke("正在检查表格结构...");
                System.Windows.Forms.Application.DoEvents();

                // 同时补全 SEQ 域信息（快速路径只扫描了光标之后的）
                if (scanStart > 1)
                {
                    try
                    {
                        Range preRange = doc.Range(tbl.Range.Start, tbl.Cell(scanStart, 1).Range.Start);
                        int preFieldCount = 0;
                        foreach (Field f in preRange.Fields)
                        {
                            try
                            {
                                if (f.Type == WdFieldType.wdFieldSequence)
                                {
                                    int rowIdx = f.Result.Cells[1].RowIndex;
                                    if (!seqFieldsByRow.ContainsKey(rowIdx))
                                        seqFieldsByRow[rowIdx] = new List<Field>();
                                    seqFieldsByRow[rowIdx].Add(f);
                                }
                            }
                            catch { }
                            // 每100个域让出UI线程
                            preFieldCount++;
                            if (preFieldCount % 100 == 0)
                                System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    catch { }
                }

                progressCallback?.Invoke(string.Format("扫描完成({0}个图片行, {1}个编号行) 已用:{2:F2}s",
                    imageRows.Count, seqFieldsByRow.Count, (DateTime.Now - startTime).TotalSeconds));
                System.Windows.Forms.Application.DoEvents();

                // === Step 3: 从第1行开始检查结构，修复增删问题 ===
                int addedCount = 0, removedCount = 0;
                bool structureChanged = false;
                int progressInterval = totalRows < 100 ? 20 : (totalRows < 500 ? 50 : 100);
        
                for (int row = 1; row <= totalRows; row++)
                {
                    // 定期更新进度，防止UI卡顿
                    if (row % progressInterval == 0)
                    {
                        progressCallback?.Invoke(string.Format("正在检查第 {0}/{1} 行...", row, totalRows));
                        System.Windows.Forms.Application.DoEvents();
                    }

                    bool hasImage = imageRows.Contains(row);
                    bool isDescRow = !hasImage && row > 1 && imageRows.Contains(row - 1);
                    bool hasSeq = seqFieldsByRow.ContainsKey(row);
        
                    if (isDescRow && !hasSeq)
                    {
                        // 缺少编号 → 插入
                        bool isFirst = false;
                        for (int col = 1; col <= colCount; col++)
                        {
                            try { InsertSeqField(tbl, row, col, wdAlignment, ref isFirst, 1); }
                            catch { }
                        }
                        addedCount++;
                        structureChanged = true;
                    }
                    else if (!isDescRow && hasSeq)
                    {
                        // 多余编号 → 移除
                        var fields = seqFieldsByRow[row];
                        for (int i = fields.Count - 1; i >= 0; i--)
                        {
                            try { fields[i].Delete(); } catch { }
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
                                    string cleaned = Regex.Replace(cellText, @"^[\.\.\s]+", "");
                                    if (cleaned != cellText) cellRange.Text = cleaned;
                                }
                            }
                            catch { }
                        }
                        seqFieldsByRow.Remove(row);
                        removedCount++;
                        structureChanged = true;
                    }
                }
        
                // === Step 4: 批量更新域值（结构变化或值不连续时） ===
                // 删除行后即使结构没变，值也可能不连续，需要强制更新
                if (structureChanged || firstBadIdx >= 0)
                {
                    progressCallback?.Invoke("正在批量更新编号值...");
                    System.Windows.Forms.Application.DoEvents();
                    try
                    {
                        // 保持 ScreenUpdating = false 进行更新，避免重绘开销
                        // 只更新表格范围内的域，而不是整个文档
                        Range tableRange = tbl.Range;
                        
                        // 如果域数量很大，分批次更新
                        int totalFields = tableRange.Fields.Count;
                        if (totalFields > 200)
                        {
                            // 大量域：逐行更新，显示进度
                            int batchSize = 50;
                            for (int i = 1; i <= totalRows; i += batchSize)
                            {
                                int endRow = Math.Min(i + batchSize - 1, totalRows);
                                try
                                {
                                    Range batchRange = doc.Range(
                                        tbl.Cell(i, 1).Range.Start,
                                        tbl.Cell(endRow, colCount).Range.End);
                                    batchRange.Fields.Update();
                                }
                                catch { }
                                
                                if (i % 100 == 0)
                                {
                                    progressCallback?.Invoke(string.Format("正在更新编号... {0}/{1} 行", 
                                        Math.Min(i + batchSize - 1, totalRows), totalRows));
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            }
                        }
                        else
                        {
                            // 少量域：直接批量更新
                            tableRange.Fields.Update();
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
                        // 让Word有机会处理重绘，但使用短延迟避免阻塞感
                        System.Windows.Forms.Application.DoEvents();
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

                // === 第二遍：编号 ===
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
                    }
                    catch { }
                }

                // 更新域：增量模式只更新 startRow 之后的范围
                try
                {
                    progressCallback?.Invoke("正在更新域...");
                    System.Windows.Forms.Application.DoEvents();

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

                // 编号完成 100%
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
        public static int CalculateNextSequenceNumber(Table tbl, int startRow)
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
        public static void InsertSeqField(Table tbl, int rowIdx, int colIdx,
            WdParagraphAlignment alignment, ref bool isFirstSeqField, int startNumber)
        {
            try
            {
                // 缓存 Cell 对象，避免重复的 tbl.Cell() 查找
                Cell cell = tbl.Cell(rowIdx, colIdx);

                // 1. 读取现有文本
                Range cellRange = cell.Range;
                cellRange.SetRange(cellRange.Start, cellRange.End - 1);
                string existingText = CleanCellText(cellRange.Text);

                // 清理已有编号前缀（如 "1. ", "2) "）
                if (!string.IsNullOrEmpty(existingText))
                {
                    existingText = Regex.Replace(existingText, @"^\d+[.)]+\s*", "");
                }

                // 2. 清空单元格内容
                cellRange.Text = "";

                // 3. 插入 SEQ 域（清空后 Range 失效，从 cell 重新获取）
                cellRange = cell.Range;
                cellRange.SetRange(cellRange.Start, cellRange.Start);

                string fieldText = (isFirstSeqField && startNumber > 1)
                    ? $"PhotoNum \\r {startNumber}"
                    : "PhotoNum";
                cellRange.Fields.Add(cellRange, WdFieldType.wdFieldSequence, fieldText, false);
                isFirstSeqField = false;

                // 4. 追加 ". " + 文件名 并设置格式（域插入后从 cell 重新获取）
                cellRange = cell.Range;
                int appendPos = cellRange.End - 1;
                cellRange.SetRange(appendPos, appendPos);

                if (!string.IsNullOrEmpty(existingText))
                {
                    cellRange.Text = ". " + existingText;
                }
                else
                {
                    cellRange.Text = ".";
                }

                // 5. 设置对齐格式
                cellRange = cell.Range;
                cellRange.SetRange(cellRange.Start, cellRange.End - 1);
                cellRange.ParagraphFormat.Alignment = alignment;
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
