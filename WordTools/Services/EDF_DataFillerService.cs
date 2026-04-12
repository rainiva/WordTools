using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordTools.Services
{
    /// <summary>
    /// Excel数据填充Word表格服务
    /// 功能：读取Excel数据并按顺序填充到Word表格
    /// </summary>
    public class EDF_DataFillerService
    {
        // 私有变量
        private string excelPath;
        private string anchorField;
        private string sampleSizeColumn;
        private bool replaceSampleSize;
        private DataTable excelData;
        private int currentRow;

        // 配置常量
        private const string CONFIG_KEY = "EDF_ExcelDataFillerConfig";

        /// <summary>
        /// 执行数据填充
        /// </summary>
        public void ExecuteFilling(string excelPath, string anchorField, int targetColumn,
            string sampleSizeColumn, bool replaceSampleSize, Action<string> onStatusUpdate)
        {
            try
            {
                // 保存参数
                this.excelPath = excelPath;
                this.anchorField = anchorField;
                this.sampleSizeColumn = sampleSizeColumn;
                this.replaceSampleSize = replaceSampleSize;

                // 验证参数
                onStatusUpdate?.Invoke("正在验证参数...");
                if (!ValidateParameters())
                {
                    onStatusUpdate?.Invoke("参数验证失败！");
                    return;
                }

                // 读取Excel数据
                onStatusUpdate?.Invoke("正在读取Excel数据...");
                if (!ReadExcelData())
                {
                    onStatusUpdate?.Invoke("读取Excel数据失败！");
                    MessageBox.Show("读取Excel数据失败！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                onStatusUpdate?.Invoke($"Excel数据读取完成，共 {excelData.Rows.Count} 行数据");
                onStatusUpdate?.Invoke("开始填充Word表格...");

                // 填充Word表格
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (FillWordTable(doc, targetColumn, onStatusUpdate))
                {
                    onStatusUpdate?.Invoke("数据填充完成！");
                    MessageBox.Show("数据填充完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    onStatusUpdate?.Invoke("数据填充过程中出现错误！");
                    MessageBox.Show("数据填充过程中出现错误！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行填充时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 确保 DataTable 资源被释放
                excelData?.Dispose();
                excelData = null;
            }
        }

        /// <summary>
        /// 验证参数
        /// </summary>
        private bool ValidateParameters()
        {
            if (string.IsNullOrEmpty(excelPath))
            {
                MessageBox.Show("请选择Excel文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!System.IO.File.Exists(excelPath))
            {
                MessageBox.Show("Excel文件不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(anchorField))
            {
                MessageBox.Show("请输入锚定字段！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 读取Excel数据
        /// </summary>
        private bool ReadExcelData()
        {
            try
            {
                string connStr;
                string extension = System.IO.Path.GetExtension(excelPath).ToLower();

                if (extension == ".xlsx")
                {
                    connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelPath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\";";
                }
                else
                {
                    connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelPath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\";";
                }

                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();

                    // 获取第一个工作表名称
                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (schemaTable == null || schemaTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Excel文件中没有工作表！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();

                    // 读取数据
                    string query = $"SELECT * FROM [{sheetName}]";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                    {
                        excelData = new DataTable();
                        adapter.Fill(excelData);
                    }

                    if (excelData.Rows.Count == 0)
                    {
                        MessageBox.Show("Excel文件中没有数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取Excel数据时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 填充Word表格
        /// </summary>
        private bool FillWordTable(Word.Document doc, int targetColumn, Action<string> onStatusUpdate)
        {
            try
            {
                bool found = false;
                int processedCount = 0;
                EDF_TemplateDetector.TemplateType lastTemplateType = EDF_TemplateDetector.TemplateType.Unknown;

                // 检查是否有表格
                if (doc.Tables.Count == 0)
                {
                    MessageBox.Show("文档中没有表格！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 保存屏幕刷新状态并关闭（提升性能）
                bool screenUpdatingState = Globals.ThisAddIn.Application.ScreenUpdating;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // 遍历所有表格
                int excelRow = 0;

                foreach (Word.Table tbl in doc.Tables)
                {
                    // 空引用检查
                    if (tbl == null)
                    {
                        Debug.WriteLine("[EDF_DataFillerService] Warning: 遇到 null 表格对象，跳过");
                        continue;
                    }

                    for (int wordRow = 1; wordRow <= tbl.Rows.Count; wordRow++)
                    {
                        // 获取锚定单元格的值
                        string anchorValue = GetCellText(tbl.Cell(wordRow, 1));

                        // 检查是否匹配锚定字段
                        if (anchorValue.IndexOf(anchorField, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // 使用模板检测器自动检测表格结构
                            var tpl = EDF_TemplateDetector.DetectTemplate(tbl, wordRow, anchorField);

                            // 如果是第一个匹配，显示检测到的模板类型
                            if (processedCount == 0)
                            {
                                onStatusUpdate?.Invoke($"检测到表格结构: {EDF_TemplateDetector.GetTemplateDescription(tpl)}");
                            }

                            // 如果模板类型变化，也显示
                            if (tpl.TemplateType != lastTemplateType)
                            {
                                onStatusUpdate?.Invoke($"切换表格结构: {EDF_TemplateDetector.GetTemplateDescription(tpl)}");
                                lastTemplateType = tpl.TemplateType;
                            }

                            // 按顺序获取下一个Excel数据行
                            excelRow++;
                            if (excelRow <= excelData.Rows.Count)
                            {
                                // 根据模板类型选择填充策略
                                switch (tpl.TemplateType)
                                {
                                    case EDF_TemplateDetector.TemplateType.StructureB:
                                        FillStructureB(tbl, wordRow, excelRow, tpl);
                                        break;
                                    case EDF_TemplateDetector.TemplateType.StructureC:
                                        FillStructureC(tbl, wordRow, excelRow, tpl);
                                        break;
                                    case EDF_TemplateDetector.TemplateType.StructureD:
                                        FillStructureD(tbl, wordRow, excelRow, tpl);
                                        break;
                                    default:
                                        FillRowData(tbl, wordRow, targetColumn, excelRow);
                                        break;
                                }

                                found = true;
                                processedCount++;

                                // 每处理50行更新一次状态
                                if (processedCount % 50 == 0)
                                {
                                    onStatusUpdate?.Invoke($"已处理 {processedCount} 行...");
                                }
                            }
                        }
                    }
                }

                // 显示处理总数
                if (found)
                {
                    onStatusUpdate?.Invoke($"填充完成，共处理 {processedCount} 行数据");
                }
                else
                {
                    // 显示诊断信息
                    string diagInfo = "未找到匹配的锚定字段。\n\n";
                    diagInfo += $"当前锚定字段设置: {anchorField}\n";
                    diagInfo += $"Excel数据行数: {excelData.Rows.Count}\n\n";

                    // 显示Word表格前5行的前两列（只显示第一个表格）
                    diagInfo += "Word表格内容（前5行）:\n";
                    if (doc.Tables.Count > 0)
                    {
                        Word.Table tbl = doc.Tables[1];
                        diagInfo += "【表格1】\n";
                        for (int wordRow = 1; wordRow <= Math.Min(tbl.Rows.Count, 5); wordRow++)
                        {
                            string col1 = GetCellText(tbl.Cell(wordRow, 1));
                            string col2 = tbl.Rows[wordRow].Cells.Count >= 2 ? GetCellText(tbl.Cell(wordRow, 2)) : "(无第二列)";
                            diagInfo += $"  行{wordRow} 列1: {col1.Substring(0, Math.Min(30, col1.Length))} | 列2: {col2.Substring(0, Math.Min(30, col2.Length))}\n";
                        }
                    }

                    // 显示Excel前5行数据
                    diagInfo += "\nExcel数据（前5行）:\n";
                    for (int i = 0; i < Math.Min(excelData.Rows.Count, 5); i++)
                    {
                        string col1 = excelData.Columns.Count >= 1 ? excelData.Rows[i][0]?.ToString() ?? "(无)" : "(无)";
                        string col2 = excelData.Columns.Count >= 2 ? excelData.Rows[i][1]?.ToString() ?? "(无)" : "(无)";
                        diagInfo += $"  行{i + 1} 列1: {col1.Substring(0, Math.Min(30, col1.Length))} | 列2: {col2.Substring(0, Math.Min(30, col2.Length))}\n";
                    }

                    MessageBox.Show(diagInfo, "锚定字段诊断", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // 恢复屏幕刷新状态
                Globals.ThisAddIn.Application.ScreenUpdating = screenUpdatingState;

                return found;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"填充Word表格时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 获取单元格文本
        /// </summary>
        private string GetCellText(Word.Cell cell)
        {
            try
            {
                string text = cell.Range.Text.Trim();
                // 移除单元格结束标记
                text = text.Replace("\r\a", "");
                return text;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in GetCellText: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 填充行数据（默认结构A）
        /// </summary>
        private void FillRowData(Word.Table tbl, int wordRow, int targetColumn, int excelRow)
        {
            try
            {
                // 确保Word表格有足够的列（添加最大列数限制防止无限循环）
                const int MAX_COLUMNS = 50;
                int addedColumns = 0;
                while (tbl.Rows[wordRow].Cells.Count < targetColumn && addedColumns < MAX_COLUMNS)
                {
                    tbl.Rows[wordRow].Cells.Add();
                    addedColumns++;
                }

                if (addedColumns >= MAX_COLUMNS)
                {
                    Debug.WriteLine($"[EDF_DataFillerService] Warning: 添加列数超过最大限制 {MAX_COLUMNS}，停止添加");
                }

                // 获取Excel第一列数据并填充到Word指定列
                if (excelData.Columns.Count >= 1)
                {
                    string dataValue = excelData.Rows[excelRow - 1][0]?.ToString() ?? "";
                    tbl.Cell(wordRow, targetColumn).Range.Text = dataValue;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in FillRowData: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充结构B（单列混合）
        /// </summary>
        private void FillStructureB(Word.Table tbl, int wordRow, int excelRow, EDF_TemplateDetector.TemplateInfo tpl)
        {
            try
            {
                string currentText = GetCellText(tbl.Cell(wordRow, 1));
                string newText = currentText;

                // 获取Excel数据
                if (excelData.Columns.Count < 1) return;
                string excelItem = (excelData.Rows[excelRow - 1][0]?.ToString() ?? "").Trim();

                // 获取Sample Size
                string excelSampleSize = "";
                if (excelData.Columns.Count >= 2)
                {
                    excelSampleSize = (excelData.Rows[excelRow - 1][1]?.ToString() ?? "").Trim();
                }

                // 1. 替换Item值
                int itemStartPos = newText.IndexOf("Item", StringComparison.OrdinalIgnoreCase);
                if (itemStartPos >= 0)
                {
                    itemStartPos = newText.IndexOf("No.", itemStartPos, StringComparison.OrdinalIgnoreCase);
                    if (itemStartPos >= 0)
                    {
                        itemStartPos += 3; // "No."的长度
                        // 跳过空格和冒号
                        while (itemStartPos < newText.Length && (newText[itemStartPos] == ' ' || newText[itemStartPos] == ':' || newText[itemStartPos] == '.'))
                        {
                            itemStartPos++;
                        }

                        // 找到Item结束位置
                        int itemEndPos = itemStartPos;
                        while (itemEndPos < newText.Length)
                        {
                            char ch = newText[itemEndPos];
                            if (ch == '(' || ch == ',' || ch == ' ') break;
                            itemEndPos++;
                        }

                        // 替换Item值
                        if (itemEndPos > itemStartPos)
                        {
                            newText = newText.Substring(0, itemStartPos) + excelItem + newText.Substring(itemEndPos);
                        }
                    }
                }

                // 2. 如果勾选了替换Sample Size，则替换
                if (replaceSampleSize && !string.IsNullOrEmpty(excelSampleSize))
                {
                    int sampleSizeStartPos = newText.IndexOf("Sample Size", StringComparison.OrdinalIgnoreCase);
                    if (sampleSizeStartPos >= 0)
                    {
                        sampleSizeStartPos = newText.IndexOf("=", sampleSizeStartPos);
                        if (sampleSizeStartPos >= 0)
                        {
                            sampleSizeStartPos++;
                            // 跳过空格
                            while (sampleSizeStartPos < newText.Length && newText[sampleSizeStartPos] == ' ')
                            {
                                sampleSizeStartPos++;
                            }

                            // 找到数字结束位置
                            int sampleSizeEndPos = sampleSizeStartPos;
                            while (sampleSizeEndPos < newText.Length && char.IsDigit(newText[sampleSizeEndPos]))
                            {
                                sampleSizeEndPos++;
                            }

                            // 替换Sample Size值
                            if (sampleSizeEndPos > sampleSizeStartPos)
                            {
                                newText = newText.Substring(0, sampleSizeStartPos) + excelSampleSize + newText.Substring(sampleSizeEndPos);
                            }
                        }
                    }
                }

                // 更新单元格内容
                tbl.Cell(wordRow, 1).Range.Text = newText;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in FillStructureB: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充结构C（多行分散）
        /// </summary>
        private void FillStructureC(Word.Table tbl, int wordRow, int excelRow, EDF_TemplateDetector.TemplateInfo tpl)
        {
            try
            {
                int targetRow = wordRow + tpl.ItemRowOffset;

                if (targetRow <= tbl.Rows.Count)
                {
                    if (excelData.Columns.Count >= 1)
                    {
                        string dataValue = excelData.Rows[excelRow - 1][0]?.ToString() ?? "";
                        tbl.Cell(targetRow, 1).Range.Text = dataValue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in FillStructureC: {ex.Message}");
            }
        }

        /// <summary>
        /// 填充结构D（合并单元格）
        /// </summary>
        private void FillStructureD(Word.Table tbl, int wordRow, int excelRow, EDF_TemplateDetector.TemplateInfo tpl)
        {
            try
            {
                if (tpl.DataColumn <= tbl.Rows[wordRow].Cells.Count)
                {
                    if (excelData.Columns.Count >= 1)
                    {
                        string dataValue = excelData.Rows[excelRow - 1][0]?.ToString() ?? "";
                        tbl.Cell(wordRow, tpl.DataColumn).Range.Text = dataValue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in FillStructureD: {ex.Message}");
            }
        }

        /// <summary>
        /// 列字母转数字
        /// </summary>
        public int ColumnLetterToNumber(string columnLetter)
        {
            int result = 0;
            columnLetter = columnLetter.Trim().ToUpper();

            for (int i = 0; i < columnLetter.Length; i++)
            {
                result = result * 26 + (columnLetter[i] - 'A' + 1);
            }

            return result;
        }

        /// <summary>
        /// 获取Excel列名（A, B, C...）
        /// </summary>
        public string GetExcelColumnName(int colIndex)
        {
            string result = "";
            while (colIndex > 0)
            {
                int remainder = (colIndex - 1) % 26;
                result = (char)('A' + remainder) + result;
                colIndex = (colIndex - 1) / 26;
            }
            return result;
        }

        /// <summary>
        /// 保存配置到文档变量
        /// </summary>
        public void SaveConfig(string key, string value)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveDocument != null)
                {
                    string fullKey = CONFIG_KEY + "_" + key;
                    // 检查变量是否存在
                    bool exists = false;
                    foreach (Word.Variable v in app.ActiveDocument.Variables)
                    {
                        if (v.Name == fullKey)
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (exists)
                    {
                        app.ActiveDocument.Variables[fullKey].Value = value;
                    }
                    else
                    {
                        app.ActiveDocument.Variables.Add(fullKey, value);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in SaveConfig: {ex.Message}");
            }
        }

        /// <summary>
        /// 加载配置
        /// </summary>
        public string LoadConfig(string key, string defaultValue)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveDocument != null)
                {
                    string fullKey = CONFIG_KEY + "_" + key;
                    foreach (Word.Variable v in app.ActiveDocument.Variables)
                    {
                        if (v.Name == fullKey)
                        {
                            return v.Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_DataFillerService] Error in LoadConfig: {ex.Message}");
            }

            return defaultValue;
        }
    }
}
