namespace WordTools.Services
{
    public struct LoggingOptionsState
    {
        public LoggingOptionsState(bool detailedLoggingEnabled, bool benchmarkLoggingEnabled)
        {
            DetailedLoggingEnabled = detailedLoggingEnabled;
            BenchmarkLoggingEnabled = benchmarkLoggingEnabled;
        }

        public bool DetailedLoggingEnabled { get; }

        public bool BenchmarkLoggingEnabled { get; }
    }

    public static class LoggingOptionsStateController
    {
        public static LoggingOptionsState Normalize(bool detailedLoggingEnabled, bool benchmarkLoggingEnabled)
        {
            return new LoggingOptionsState(
                detailedLoggingEnabled,
                detailedLoggingEnabled && benchmarkLoggingEnabled);
        }

        public static bool ShouldShowDetailedLog(bool detailedLoggingEnabled)
        {
            return detailedLoggingEnabled;
        }

        public static bool ShouldWriteBenchmarkLog(bool detailedLoggingEnabled, bool benchmarkLoggingEnabled)
        {
            return detailedLoggingEnabled && benchmarkLoggingEnabled;
        }
    }
}
