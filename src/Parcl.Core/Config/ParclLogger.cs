using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

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
        private readonly ConcurrentQueue<string> _buffer = new ConcurrentQueue<string>();
        private readonly Timer _flushTimer;
        private readonly string _logFile;
        private bool _disposed;

        public ParclLogger(LogLevel minLevel = LogLevel.Debug)
        {
            _minLevel = minLevel;

            Directory.CreateDirectory(LogDir);
            CleanOldLogs(maxAgeDays: 7);

            _logFile = Path.Combine(LogDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");
            _writer = new StreamWriter(_logFile, append: true, Encoding.UTF8)
            {
                AutoFlush = false
            };

            // Flush every 2 seconds to avoid I/O thrashing
            _flushTimer = new Timer(_ => Flush(), null, 2000, 2000);

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

            // Format: 2026-03-16T14:30:45.123 [INFO ] [LDAP    ] Found 2 certs for user@example.com
            var timestamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff", CultureInfo.InvariantCulture);
            var lvl = level.ToString().ToUpper().PadRight(5);
            var comp = component.PadRight(8);
            var line = $"{timestamp} [{lvl}] [{comp}] {message}";

            _buffer.Enqueue(line);
        }

        private void Flush()
        {
            if (_disposed) return;

            try
            {
                while (_buffer.TryDequeue(out var line))
                {
                    _writer.WriteLine(line);
                }
                _writer.Flush();
            }
            catch
            {
                // Can't log a logging failure — just drop it
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
                    {
                        File.Delete(file);
                    }
                }
            }
            catch
            {
                // Non-critical — old logs stay a bit longer
            }
        }

        /// <summary>
        /// Returns the path to today's log file for user troubleshooting.
        /// </summary>
        public string GetLogFilePath() => _logFile;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            Info("Logger", "Session ending — flushing logs");
            _flushTimer?.Dispose();
            Flush();
            _writer?.Dispose();
        }
    }
}
