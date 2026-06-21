using System;
using System.Collections.Generic;
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
                        if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\s*\S"))
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

        public static int GetTableRowCount(Document doc)
        {
            if (doc == null || doc.Tables.Count == 0)
            {
                return 0;
            }

            return doc.Tables[1].Rows.Count;
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
