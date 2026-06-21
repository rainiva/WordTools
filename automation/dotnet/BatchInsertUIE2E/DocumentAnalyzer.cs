using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace BatchInsertUIE2E
{
    internal static class DocumentAnalyzer
    {
        public static bool HasNumberedDescription(Document doc)
        {
            if (doc.Tables.Count == 0)
            {
                return false;
            }

            var table = doc.Tables[1];
            var numberedCount = 0;
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                for (var col = 1; col <= 2; col++)
                {
                    try
                    {
                        var text = NormalizeCellText(table.Cell(row, col).Range.Text);
                        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\."))
                        {
                            numberedCount++;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return numberedCount >= 2;
        }

        public static bool HasSubfolderTitle(Document doc, string subfolderName)
        {
            if (doc.Tables.Count == 0 || string.IsNullOrWhiteSpace(subfolderName))
            {
                return false;
            }

            var table = doc.Tables[1];
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                try
                {
                    var text = NormalizeCellText(table.Cell(row, 1).Range.Text);
                    if (text.IndexOf(subfolderName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        public static bool HasAnySubfolderTitle(Document doc, IEnumerable<string> subfolderNames)
        {
            if (subfolderNames == null)
            {
                return false;
            }

            return subfolderNames.Any(name => HasSubfolderTitle(doc, name));
        }

        public static string GetLastImageRowCol2Text(Document doc)
        {
            if (doc.Tables.Count == 0 || doc.InlineShapes.Count == 0)
            {
                return string.Empty;
            }

            var table = doc.Tables[1];
            for (var row = table.Rows.Count; row >= 1; row--)
            {
                try
                {
                    var cell = table.Cell(row, 1).Range;
                    if (cell.InlineShapes.Count > 0)
                    {
                        return NormalizeCellText(table.Cell(row, 2).Range.Text);
                    }
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        public static bool HasNumberAfterDescription(Document doc)
        {
            return FindDescriptionCell(doc, text => System.Text.RegularExpressions.Regex.IsMatch(text, @"-\d+$"));
        }

        public static bool HasCenterAlignedNumberedDescription(Document doc)
        {
            return HasAlignedNumberedDescription(doc, WdParagraphAlignment.wdAlignParagraphCenter);
        }

        public static bool HasLeftAlignedNumberedDescription(Document doc)
        {
            return HasAlignedNumberedDescription(doc, WdParagraphAlignment.wdAlignParagraphLeft);
        }

        public static bool HasFolderNameDescription(Document doc, IEnumerable<string> subfolderNames)
        {
            if (doc.Tables.Count == 0 || subfolderNames == null)
            {
                return false;
            }

            var table = doc.Tables[1];
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                for (var col = 1; col <= 2; col++)
                {
                    try
                    {
                        var text = NormalizeCellText(table.Cell(row, col).Range.Text);
                        if (string.IsNullOrWhiteSpace(text) || text.IndexOf(".jpg", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            continue;
                        }

                        if (subfolderNames.Any(name => text.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return false;
        }

        public static bool HasManualDescriptionRows(Document doc)
        {
            if (doc.Tables.Count == 0)
            {
                return false;
            }

            var table = doc.Tables[1];
            var manualRows = 0;
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                try
                {
                    if (table.Cell(row, 1).Range.InlineShapes.Count > 0)
                    {
                        continue;
                    }

                    var text = NormalizeCellText(table.Cell(row, 1).Range.Text);
                    if (string.IsNullOrWhiteSpace(text) || System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.?$"))
                    {
                        manualRows++;
                    }
                }
                catch
                {
                }
            }

            return manualRows >= 2;
        }

        public static int GetTableRowCount(Document doc)
        {
            if (doc == null || doc.Tables.Count == 0)
            {
                return 0;
            }

            return doc.Tables[1].Rows.Count;
        }

        private static bool HasAlignedNumberedDescription(Document doc, WdParagraphAlignment expected)
        {
            if (doc.Tables.Count == 0)
            {
                return false;
            }

            var table = doc.Tables[1];
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                for (var col = 1; col <= 2; col++)
                {
                    try
                    {
                        var cell = table.Cell(row, col);
                        var text = NormalizeCellText(cell.Range.Text);
                        if (!System.Text.RegularExpressions.Regex.IsMatch(text, @"\d"))
                        {
                            continue;
                        }

                        if (cell.Range.ParagraphFormat.Alignment == expected)
                        {
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return false;
        }

        private static bool FindDescriptionCell(Document doc, Func<string, bool> predicate)
        {
            if (doc.Tables.Count == 0)
            {
                return false;
            }

            var table = doc.Tables[1];
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                for (var col = 1; col <= 2; col++)
                {
                    try
                    {
                        var text = NormalizeCellText(table.Cell(row, col).Range.Text);
                        if (predicate(text))
                        {
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return false;
        }

        private static string NormalizeCellText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text.Replace("\r", string.Empty)
                .Replace("\n", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\x01", string.Empty)
                .Trim();
        }
    }
}
