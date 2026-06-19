using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace WordTools.Services
{
    public sealed class BenchmarkLogEntry
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;
        public string RunMode { get; set; }
        public string Status { get; set; }
        public string DocumentPath { get; set; }
        public string LogPath { get; set; }
        public string SourcePath { get; set; }
        public int TotalFiles { get; set; }
        public int ProcessedCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public int MergedCellCount { get; set; }
        public bool Cancelled { get; set; }
        public bool NeedDescription { get; set; }
        public bool UseFileNameAsDescription { get; set; }
        public bool UseFolderNameAsDescription { get; set; }
        public bool AutoNumbering { get; set; }
        public int NumberAlignment { get; set; }
        public int NumberPosition { get; set; }
        public double TotalSeconds { get; set; }
        public long? InitializeMs { get; set; }
        public long? ClearNumberingMs { get; set; }
        public long? CalculateStartNumberMs { get; set; }
        public long? PreAllocateRowsMs { get; set; }
        public long? InsertImagesMs { get; set; }
        public long? WrapUpMs { get; set; }
        public long? CellAvailabilityMs { get; set; }
        public int? CellAvailabilityCount { get; set; }
        public long? FloatingShapeLookupMs { get; set; }
        public int? FloatingShapeLookupCount { get; set; }
        public long? OverwriteClearMs { get; set; }
        public int? OverwriteClearCount { get; set; }
        public long? AddPictureMs { get; set; }
        public int? AddPictureCount { get; set; }
        public long? CellValidationMs { get; set; }
        public int? CellValidationCount { get; set; }
        public long? PictureSizingMs { get; set; }
        public int? PictureSizingCount { get; set; }
        public long? ProgressUiMs { get; set; }
        public int? ProgressUiCount { get; set; }
        public long? DescriptionWriteMs { get; set; }
        public int? DescriptionWriteCount { get; set; }
        public bool? SkippedClearNumbering { get; set; }
        public string ErrorMessage { get; set; }
    }

    public static class BenchmarkLogService
    {
        private const string FileName = "wordtools-benchmark.csv";
        private const string Header =
            "timestamp_utc,run_mode,status,document_path,log_path,source_path,total_files,processed_count,success_count,fail_count,merged_cell_count,cancelled,need_description,use_filename_as_description,use_foldername_as_description,auto_numbering,number_alignment,number_position,total_seconds,initialize_ms,clear_numbering_ms,calculate_start_number_ms,preallocate_rows_ms,insert_images_ms,wrap_up_ms,cell_availability_ms,cell_availability_count,floating_shape_lookup_ms,floating_shape_lookup_count,overwrite_clear_ms,overwrite_clear_count,add_picture_ms,add_picture_count,cell_validation_ms,cell_validation_count,picture_sizing_ms,picture_sizing_count,progress_ui_ms,progress_ui_count,description_write_ms,description_write_count,skipped_clear_numbering,error_message";

        public static string GetDefaultLogPath(string documentPath)
        {
            string baseDirectory = null;

            if (!string.IsNullOrWhiteSpace(documentPath))
            {
                try
                {
                    baseDirectory = Path.GetDirectoryName(documentPath);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[BenchmarkLogService] GetDefaultLogPath error: {ex.Message}");
                    baseDirectory = null;
                }
            }

            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                baseDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "WordTools");
            }

            return Path.Combine(baseDirectory, FileName);
        }

        public static void AppendCsv(string filePath, BenchmarkLogEntry entry)
        {
            if (entry == null)
            {
                return;
            }

            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            bool writeHeader = !File.Exists(filePath);
            using (var stream = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                if (writeHeader)
                {
                    writer.WriteLine(Header);
                }

                writer.WriteLine(BuildCsvLine(entry));
            }
        }

        private static string BuildCsvLine(BenchmarkLogEntry entry)
        {
            var values = new[]
            {
                entry.TimestampUtc.ToString("O", CultureInfo.InvariantCulture),
                entry.RunMode,
                entry.Status,
                entry.DocumentPath,
                entry.LogPath,
                entry.SourcePath,
                entry.TotalFiles.ToString(CultureInfo.InvariantCulture),
                entry.ProcessedCount.ToString(CultureInfo.InvariantCulture),
                entry.SuccessCount.ToString(CultureInfo.InvariantCulture),
                entry.FailCount.ToString(CultureInfo.InvariantCulture),
                entry.MergedCellCount.ToString(CultureInfo.InvariantCulture),
                entry.Cancelled ? "True" : "False",
                entry.NeedDescription ? "True" : "False",
                entry.UseFileNameAsDescription ? "True" : "False",
                entry.UseFolderNameAsDescription ? "True" : "False",
                entry.AutoNumbering ? "True" : "False",
                entry.NumberAlignment.ToString(CultureInfo.InvariantCulture),
                entry.NumberPosition.ToString(CultureInfo.InvariantCulture),
                entry.TotalSeconds.ToString("F3", CultureInfo.InvariantCulture),
                FormatNullableLong(entry.InitializeMs),
                FormatNullableLong(entry.ClearNumberingMs),
                FormatNullableLong(entry.CalculateStartNumberMs),
                FormatNullableLong(entry.PreAllocateRowsMs),
                FormatNullableLong(entry.InsertImagesMs),
                FormatNullableLong(entry.WrapUpMs),
                FormatNullableLong(entry.CellAvailabilityMs),
                FormatNullableInt(entry.CellAvailabilityCount),
                FormatNullableLong(entry.FloatingShapeLookupMs),
                FormatNullableInt(entry.FloatingShapeLookupCount),
                FormatNullableLong(entry.OverwriteClearMs),
                FormatNullableInt(entry.OverwriteClearCount),
                FormatNullableLong(entry.AddPictureMs),
                FormatNullableInt(entry.AddPictureCount),
                FormatNullableLong(entry.CellValidationMs),
                FormatNullableInt(entry.CellValidationCount),
                FormatNullableLong(entry.PictureSizingMs),
                FormatNullableInt(entry.PictureSizingCount),
                FormatNullableLong(entry.ProgressUiMs),
                FormatNullableInt(entry.ProgressUiCount),
                FormatNullableLong(entry.DescriptionWriteMs),
                FormatNullableInt(entry.DescriptionWriteCount),
                entry.SkippedClearNumbering.HasValue ? (entry.SkippedClearNumbering.Value ? "True" : "False") : "",
                entry.ErrorMessage
            };

            return string.Join(",", EscapeValues(values));
        }

        private static IEnumerable<string> EscapeValues(IEnumerable<string> values)
        {
            foreach (string raw in values)
            {
                yield return EscapeCsv(raw);
            }
        }

        private static string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            string normalized = value
                .Replace("\r\n", " ")
                .Replace("\r", " ")
                .Replace("\n", " ");
            while (normalized.Contains("  "))
            {
                normalized = normalized.Replace("  ", " ");
            }
            bool needsQuotes = normalized.IndexOfAny(new[] { ',', '"' }) >= 0;
            if (normalized.Contains("\""))
            {
                normalized = normalized.Replace("\"", "\"\"");
                needsQuotes = true;
            }

            return needsQuotes ? "\"" + normalized + "\"" : normalized;
        }

        private static string FormatNullableLong(long? value)
        {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : "";
        }

        private static string FormatNullableInt(int? value)
        {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : "";
        }
    }
}
