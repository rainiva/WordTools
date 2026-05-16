using System.Text;

namespace WordTools.Services
{
    public sealed class InsertionPerformanceDiagnostics
    {
        public long CellAvailabilityMs { get; private set; }
        public int CellAvailabilityCount { get; private set; }

        public long FloatingShapeLookupMs { get; private set; }
        public int FloatingShapeLookupCount { get; private set; }

        public long OverwriteClearMs { get; private set; }
        public int OverwriteClearCount { get; private set; }

        public long AddPictureMs { get; private set; }
        public int AddPictureCount { get; private set; }

        public long CellValidationMs { get; private set; }
        public int CellValidationCount { get; private set; }

        public long PictureSizingMs { get; private set; }
        public int PictureSizingCount { get; private set; }

        public long ProgressUiMs { get; private set; }
        public int ProgressUiCount { get; private set; }

        public long DescriptionWriteMs { get; private set; }
        public int DescriptionWriteCount { get; private set; }

        public void RecordCellAvailabilityCheck(long elapsedMs)
        {
            CellAvailabilityMs += Normalize(elapsedMs);
            CellAvailabilityCount++;
        }

        public void RecordFloatingShapeLookup(long elapsedMs)
        {
            FloatingShapeLookupMs += Normalize(elapsedMs);
            FloatingShapeLookupCount++;
        }

        public void RecordOverwriteClear(long elapsedMs)
        {
            OverwriteClearMs += Normalize(elapsedMs);
            OverwriteClearCount++;
        }

        public void RecordAddPicture(long elapsedMs)
        {
            AddPictureMs += Normalize(elapsedMs);
            AddPictureCount++;
        }

        public void RecordCellValidation(long elapsedMs)
        {
            CellValidationMs += Normalize(elapsedMs);
            CellValidationCount++;
        }

        public void RecordPictureSizing(long elapsedMs)
        {
            PictureSizingMs += Normalize(elapsedMs);
            PictureSizingCount++;
        }

        public void RecordProgressUi(long elapsedMs)
        {
            ProgressUiMs += Normalize(elapsedMs);
            ProgressUiCount++;
        }

        public void RecordDescriptionWrite(long elapsedMs)
        {
            DescriptionWriteMs += Normalize(elapsedMs);
            DescriptionWriteCount++;
        }

        public string BuildDetailedLog()
        {
            var sb = new StringBuilder();
            sb.AppendLine(BuildLine("行可用性检查", CellAvailabilityMs, CellAvailabilityCount));
            sb.AppendLine(BuildLine("浮动图片探测", FloatingShapeLookupMs, FloatingShapeLookupCount));
            sb.AppendLine(BuildLine("覆盖清理", OverwriteClearMs, OverwriteClearCount));
            sb.AppendLine(BuildLine("AddPicture", AddPictureMs, AddPictureCount));
            sb.AppendLine(BuildLine("Cell validation", CellValidationMs, CellValidationCount));
            sb.AppendLine(BuildLine("Picture sizing", PictureSizingMs, PictureSizingCount));
            sb.AppendLine(BuildLine("Progress UI", ProgressUiMs, ProgressUiCount));
            sb.Append(BuildLine("描述/编号写入", DescriptionWriteMs, DescriptionWriteCount));
            return sb.ToString();
        }

        private static long Normalize(long elapsedMs)
        {
            return elapsedMs < 0 ? 0 : elapsedMs;
        }

        private static string BuildLine(string label, long totalMs, int count)
        {
            return string.Format("{0}: {1}ms ({2}次)", label, totalMs, count);
        }
    }
}
