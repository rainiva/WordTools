namespace WordTools.Services.Abstractions
{
    /// <summary>性能基准 CSV 日志抽象。</summary>
    /// <remarks>规划抽象（Phase 2）。运行时仍使用 static BenchmarkLogService，尚无 Adapter 实现。</remarks>
    public interface IBenchmarkLogService
    {
        string GetDefaultLogPath(string documentPath);
        void AppendCsv(string filePath, BenchmarkLogEntry entry);
    }
}
