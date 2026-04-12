using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordTools.Services
{
    /// <summary>
    /// 智能模板检测器
    /// 功能：自动检测Word表格结构并选择合适的填充策略
    /// </summary>
    public static class EDF_TemplateDetector
    {
        /// <summary>
        /// 模板类型枚举
        /// </summary>
        public enum TemplateType
        {
            Unknown = 0,
            StructureA = 1,      // 两列分离
            StructureB = 2,      // 单列混合
            StructureC = 3,      // 多行分散
            StructureD = 4       // 合并单元格
        }

        /// <summary>
        /// 模板信息结构
        /// </summary>
        public class TemplateInfo
        {
            public TemplateType TemplateType { get; set; }
            public int ItemColumn { get; set; }          // Item值所在列
            public int DataColumn { get; set; }          // 数据填充目标列
            public int ItemRowOffset { get; set; }       // Item值相对锚定行的行偏移

            public TemplateInfo()
            {
                TemplateType = TemplateType.Unknown;
                ItemColumn = 2;
                DataColumn = 2;
                ItemRowOffset = 0;
            }
        }

        /// <summary>
        /// 获取单元格文本（本地副本）
        /// </summary>
        private static string GetCellText(Word.Cell cell)
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
                Debug.WriteLine($"[EDF_TemplateDetector] Error: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 主检测函数
        /// </summary>
        public static TemplateInfo DetectTemplate(Word.Table tbl, int anchorRow, string anchorField)
        {
            // 空引用检查
            if (tbl == null)
            {
                return new TemplateInfo();
            }

            TemplateInfo result = new TemplateInfo();

            // 获取锚定单元格内容
            string anchorText = GetCellText(tbl.Cell(anchorRow, 1));

            // 检测各结构特征
            if (IsStructureB(tbl, anchorRow, anchorText))
            {
                result.TemplateType = TemplateType.StructureB;
                result.ItemColumn = 1;
                result.DataColumn = 1;
            }
            else if (IsStructureC(tbl, anchorRow, anchorText))
            {
                result.TemplateType = TemplateType.StructureC;
                result.ItemColumn = 1;
                result.DataColumn = 1;
                result.ItemRowOffset = 1;
            }
            else if (IsStructureD(tbl, anchorRow, anchorText))
            {
                result.TemplateType = TemplateType.StructureD;
                result.ItemColumn = 2;
                result.DataColumn = 2;
            }
            else if (IsStructureA(tbl, anchorRow, anchorText))
            {
                result.TemplateType = TemplateType.StructureA;
                result.ItemColumn = 2;
                result.DataColumn = 2;
            }

            return result;
        }

        /// <summary>
        /// 结构A检测：两列分离
        /// </summary>
        private static bool IsStructureA(Word.Table tbl, int row, string cellText)
        {
            // 检查是否有第二列
            if (tbl.Rows[row].Cells.Count < 2)
            {
                return false;
            }

            // 检查第二列是否有内容
            string col2Text = GetCellText(tbl.Cell(row, 2)).Trim();

            // 第二列有内容且不包含"Item"关键字
            if (!string.IsNullOrEmpty(col2Text) && col2Text.IndexOf("Item", StringComparison.OrdinalIgnoreCase) < 0)
            {
                // 检查内容是否像Item编号（包含字母数字和连字符）
                if (IsValidItemCode(col2Text))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 结构B检测：单列混合
        /// </summary>
        private static bool IsStructureB(Word.Table tbl, int row, string cellText)
        {
            // 检查是否包含逗号或等号（混合格式特征）
            if (cellText.Contains(",") || cellText.Contains("="))
            {
                // 检查是否包含Item No.和可能的Sample Size
                if (cellText.IndexOf("Item", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 检查是否在同一行包含Item值（通过正则）
                    if (!string.IsNullOrEmpty(ExtractItemByRegex(cellText)))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 结构C检测：多行分散
        /// </summary>
        private static bool IsStructureC(Word.Table tbl, int row, string cellText)
        {
            // 检查当前行第二列是否为空
            if (tbl.Rows[row].Cells.Count >= 2)
            {
                string col2Text = GetCellText(tbl.Cell(row, 2)).Trim();
                if (!string.IsNullOrEmpty(col2Text))
                {
                    return false;
                }
            }

            // 检查下一行是否存在且第一列看起来像Item值
            // Word表格行索引从1开始，所以检查 row + 1 <= tbl.Rows.Count
            if (row + 1 <= tbl.Rows.Count)
            {
                string nextRowText = GetCellText(tbl.Cell(row + 1, 1)).Trim();

                // 下一行第一列是Item值（不包含Item关键字，但像Item编号）
                if (nextRowText.IndexOf("Item", StringComparison.OrdinalIgnoreCase) < 0 && IsValidItemCode(nextRowText))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 结构D检测：合并单元格
        /// </summary>
        private static bool IsStructureD(Word.Table tbl, int row, string cellText)
        {
            // 检查第一列是否跨多列合并
            if (tbl.Rows[row].Cells.Count > 1)
            {
                // 检查第一列的宽度是否异常（合并单元格特征）
                string col2Text = GetCellText(tbl.Cell(row, 2)).Trim();

                // 如果第二列看起来是空的，可能是合并单元格
                if (string.IsNullOrEmpty(col2Text) || col2Text == "\r\a")
                {
                    // 再检查是否有第三列或更多列
                    if (tbl.Rows[row].Cells.Count >= 3)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 使用正则表达式提取Item值
        /// </summary>
        public static string ExtractItemByRegex(string cellText)
        {
            try
            {
                // 模式1：Item No. 后面跟着的值（到逗号、空格或左括号为止）
                Regex regex1 = new Regex(@"Item\s*No\.?\s*:?\s*([A-Z0-9\-]+)", RegexOptions.IgnoreCase);
                Match match1 = regex1.Match(cellText);
                if (match1.Success && match1.Groups.Count > 1)
                {
                    return match1.Groups[1].Value;
                }

                // 模式2：查找任何看起来像Item编号的格式（2-3个字母开头）
                Regex regex2 = new Regex(@"\b([A-Z]{2,3}-[A-Z0-9\-]+)\b");
                Match match2 = regex2.Match(cellText);
                if (match2.Success && match2.Groups.Count > 1)
                {
                    return match2.Groups[1].Value;
                }

                // 模式3：查找纯数字或字母数字组合
                Regex regex3 = new Regex(@"\b([A-Z0-9]{3,})\b");
                Match match3 = regex3.Match(cellText);
                if (match3.Success && match3.Groups.Count > 1)
                {
                    return match3.Groups[1].Value;
                }

                return "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_TemplateDetector] Error: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 验证是否为有效的Item编号
        /// </summary>
        private static bool IsValidItemCode(string text)
        {
            try
            {
                // Item编号通常包含：字母、数字、连字符
                Regex regex = new Regex(@"^[A-Z0-9\-]+$", RegexOptions.IgnoreCase);
                return regex.IsMatch(text.Trim());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[EDF_TemplateDetector] Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取模板描述文本
        /// </summary>
        public static string GetTemplateDescription(TemplateInfo tpl)
        {
            switch (tpl.TemplateType)
            {
                case TemplateType.StructureA:
                    return "结构A-两列分离";
                case TemplateType.StructureB:
                    return "结构B-单列混合";
                case TemplateType.StructureC:
                    return "结构C-多行分散";
                case TemplateType.StructureD:
                    return "结构D-合并单元格";
                default:
                    return "未知结构";
            }
        }
    }
}
