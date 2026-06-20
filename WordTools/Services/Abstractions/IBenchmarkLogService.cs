namespace WordTools.Services.Abstractions
{
    public interface IBenchmarkLogService
    {
        string GetDefaultLogPath(string documentPath);
        void AppendCsv(string filePath, BenchmarkLogEntry entry);
    }
}
