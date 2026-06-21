using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace BatchInsertE2E
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
            if (doc.Tables.Count == 0)
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
                    // merged cells may throw on column 2; title row is merged across columns
                }
            }

            return false;
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

        public static string[] CollectDescriptionSamples(Document doc)
        {
            var samples = new List<string>();
            if (doc.Tables.Count == 0)
            {
                return samples.ToArray();
            }

            var table = doc.Tables[1];
            for (var row = 1; row <= table.Rows.Count; row++)
            {
                var text = NormalizeCellText(table.Cell(row, 1).Range.Text);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    samples.Add(text);
                }
            }

            return samples.ToArray();
        }

        public static bool HasNumberAfterDescription(Document doc)
        {
            return FindNumberedDescriptionCell(doc, text =>
                System.Text.RegularExpressions.Regex.IsMatch(text, @"-\d+$"));
        }

        public static bool HasCenterAlignedNumberedDescription(Document doc)
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

                        var alignment = cell.Range.ParagraphFormat.Alignment;
                        if (alignment == WdParagraphAlignment.wdAlignParagraphCenter)
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

        public static bool HasFolderNameDescription(Document doc)
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
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        if (text.IndexOf("sub-a", StringComparison.OrdinalIgnoreCase) >= 0
                            && text.IndexOf(".jpg", StringComparison.OrdinalIgnoreCase) < 0)
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
                    var imageCell = table.Cell(row, 1).Range;
                    if (imageCell.InlineShapes.Count > 0)
                    {
                        continue;
                    }

                    var text = NormalizeCellText(table.Cell(row, 1).Range.Text);
                    if (string.IsNullOrWhiteSpace(text))
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

        private static bool FindNumberedDescriptionCell(Document doc, Func<string, bool> predicate)
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
