using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace Parcl.Core.Config
{
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warn = 2,
        Error = 3
    }

    public class ParclLogger : IDisposable
    {
        private static readonly string LogDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl", "logs");

        private readonly LogLevel _minLevel;
        private readonly StreamWriter _writer;
        private readonly string _logFile;
        private readonly object _lock = new object();
        private bool _disposed;

        public ParclLogger(LogLevel minLevel = LogLevel.Debug)
        {
            _minLevel = minLevel;

            Directory.CreateDirectory(LogDir);
            CleanOldLogs(maxAgeDays: 7);

            _logFile = Path.Combine(LogDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");
            _writer = new StreamWriter(
                new FileStream(_logFile, FileMode.Append, FileAccess.Write, FileShare.Read),
                Encoding.UTF8);

            Info("Logger", $"Session started — level={_minLevel}, pid={System.Diagnostics.Process.GetCurrentProcess().Id}");
        }

        /// <summary>
        /// Operational detail useful during development or diagnosing specific issues.
        /// Use for: method entry/exit with key params, cache hits/misses, config values loaded.
        /// Do NOT use for: loop iterations, per-byte processing, UI redraws.
        /// </summary>
        public void Debug(string component, string message)
            => Write(LogLevel.Debug, component, message);

        /// <summary>
        /// Significant operational events that confirm the system is working correctly.
        /// Use for: user actions (encrypt, sign, lookup), connection established, cert loaded.
        /// Keep these meaningful — each INFO line should tell you something happened.
        /// </summary>
        public void Info(string component, string message)
            => Write(LogLevel.Info, component, message);

        /// <summary>
        /// Something unexpected that the system recovered from, but deserves attention.
        /// Use for: cert expiring soon, LDAP fallback to next server, deprecated config found.
        /// </summary>
        public void Warn(string component, string message)
            => Write(LogLevel.Warn, component, message);

        /// <summary>
        /// Operation failed. The user's action did not complete as expected.
        /// Use for: LDAP connection failure, decrypt failed, cert store access denied.
        /// Always include the actionable detail — what failed, what was attempted, what to check.
        /// </summary>
        public void Error(string component, string message)
            => Write(LogLevel.Error, component, message);

        /// <summary>
        /// Error with exception details. Captures type, message, and first stack frame.
        /// </summary>
        public void Error(string component, string message, Exception ex)
        {
            var firstFrame = ex.StackTrace?.Split('\n')[0]?.Trim() ?? "no stack";
            Write(LogLevel.Error, component, $"{message} | {ex.GetType().Name}: {ex.Message} | at {firstFrame}");
        }

        private void Write(LogLevel level, string component, string message)
        {
            if (level < _minLevel || _disposed) return;

            var timestamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture);
            var lvl = level.ToString().ToUpper().PadRight(5);
            var comp = component.PadRight(8);
            var line = $"{timestamp} [{lvl}] [{comp}] {message}";

            lock (_lock)
            {
                if (_disposed) return;
                _writer.WriteLine(line);
                _writer.Flush();
            }
        }

        private static void CleanOldLogs(int maxAgeDays)
        {
            try
            {
                var cutoff = DateTime.Now.AddDays(-maxAgeDays);
                foreach (var file in Directory.GetFiles(LogDir, "parcl-*.log"))
                {
                    if (File.GetCreationTime(file) < cutoff)
                        File.Delete(file);
                }
            }
            catch { }
        }

        public string GetLogFilePath() => _logFile;

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;

                // Write final message directly (bypassing _disposed check)
                var timestamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture);
                _writer.WriteLine($"{timestamp} [INFO ] [Logger  ] Session ending");

                _disposed = true;
                _writer.Flush();
                _writer.Dispose();
            }
        }
    }
}
