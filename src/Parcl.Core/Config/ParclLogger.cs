using System;
using System.Diagnostics;
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

    /// <summary>
    /// Structured JSONL logger. Each line is a self-contained JSON object.
    ///
    /// Filter examples:
    ///   PowerShell:  Get-Content parcl.jsonl | ConvertFrom-Json | Where-Object cmp -eq "LDAP"
    ///   PowerShell:  Get-Content parcl.jsonl | ConvertFrom-Json | Where-Object lvl -eq "ERROR" | Format-Table
    ///   jq:          jq 'select(.cmp=="Encrypt")' parcl.jsonl
    ///   Excel:       Data → Get Data → From JSON → select the .jsonl file
    /// </summary>
    public class ParclLogger : IDisposable
    {
        private static readonly string DefaultLogDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl", "logs");

        private readonly string _logDir;
        private readonly string _sessionId;
        private readonly int _pid;
        private LogLevel _minLevel;
        private readonly StreamWriter _writer;
        private readonly string _logFile;
        private readonly object _lock = new object();
        private bool _disposed;

        public ParclLogger(LogLevel minLevel = LogLevel.Debug, string? logDirectory = null)
        {
            _minLevel = minLevel;
            _logDir = logDirectory ?? DefaultLogDir;
            _pid = Process.GetCurrentProcess().Id;
            _sessionId = GenerateSessionId();

            Directory.CreateDirectory(_logDir);
            CleanOldLogs(maxAgeDays: 7);

            _logFile = Path.Combine(_logDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.jsonl");
            _writer = new StreamWriter(
                new FileStream(_logFile, FileMode.Append, FileAccess.Write, FileShare.Read),
                Encoding.UTF8)
            { AutoFlush = true };

            Info("Logger", $"Session started — level={_minLevel}");
        }

        public void Debug(string component, string message)
            => Write(LogLevel.Debug, component, message);

        public void Info(string component, string message)
            => Write(LogLevel.Info, component, message);

        public void Warn(string component, string message)
            => Write(LogLevel.Warn, component, message);

        public void Error(string component, string message)
            => Write(LogLevel.Error, component, message);

        public void Error(string component, string message, Exception ex)
        {
            var firstFrame = ex.StackTrace?.Split('\n')[0]?.Trim() ?? "no stack";
            WriteError(component, message, ex.GetType().Name, ex.Message, firstFrame);
        }

        private void Write(LogLevel level, string component, string message)
        {
            if (level < _minLevel || _disposed) return;

            var ts = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
            var lvl = level.ToString().ToUpper();
            var line = $"{{\"ts\":\"{ts}\",\"lvl\":\"{lvl}\",\"cmp\":\"{EscapeJson(component)}\"," +
                       $"\"sid\":\"{_sessionId}\",\"pid\":{_pid}," +
                       $"\"msg\":\"{EscapeJson(message)}\"}}";

            WriteLine(line);
        }

        private void WriteError(string component, string message,
            string exType, string exMsg, string frame)
        {
            if (_disposed) return;

            var ts = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
            var line = $"{{\"ts\":\"{ts}\",\"lvl\":\"ERROR\",\"cmp\":\"{EscapeJson(component)}\"," +
                       $"\"sid\":\"{_sessionId}\",\"pid\":{_pid}," +
                       $"\"msg\":\"{EscapeJson(message)}\"," +
                       $"\"err\":{{\"type\":\"{EscapeJson(exType)}\",\"msg\":\"{EscapeJson(exMsg)}\"," +
                       $"\"at\":\"{EscapeJson(frame)}\"}}}}";

            WriteLine(line);
        }

        private void WriteLine(string line)
        {
            lock (_lock)
            {
                if (_disposed) return;
                _writer.WriteLine(line);
                _writer.Flush();
            }
        }

        private void CleanOldLogs(int maxAgeDays)
        {
            try
            {
                var cutoff = DateTime.Now.AddDays(-maxAgeDays);
                foreach (var file in Directory.GetFiles(_logDir, "parcl-*.jsonl"))
                {
                    if (File.GetLastWriteTime(file) < cutoff)
                        File.Delete(file);
                }
                // Also clean legacy .log files
                foreach (var file in Directory.GetFiles(_logDir, "parcl-*.log"))
                {
                    if (File.GetLastWriteTime(file) < cutoff)
                        File.Delete(file);
                }
            }
            catch { }
        }

        public string GetLogFilePath() => _logFile;
        public string GetLogDirectory() => _logDir;

        public void SetMinLevel(LogLevel level)
        {
            var old = _minLevel;
            _minLevel = level;
            Info("Logger", $"Log level changed: {old} -> {level}");
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;

                var ts = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
                _writer.WriteLine(
                    $"{{\"ts\":\"{ts}\",\"lvl\":\"INFO\",\"cmp\":\"Logger\"," +
                    $"\"sid\":\"{_sessionId}\",\"pid\":{_pid}," +
                    $"\"msg\":\"Session ending\"}}");

                _disposed = true;
                _writer.Flush();
                _writer.Dispose();
            }
        }

        private static string GenerateSessionId()
        {
            // Short 6-char hex ID — enough to distinguish concurrent sessions
            var bytes = new byte[3];
            using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
                rng.GetBytes(bytes);
            return BitConverter.ToString(bytes).Replace("-", "").ToLower();
        }

        /// <summary>
        /// Masks the local part of an email address for PII-safe logging.
        /// "user@quantumnexum.com" becomes "us***@quantumnexum.com".
        /// Returns the original string unchanged if it does not contain '@'.
        /// </summary>
        public static string SanitizeEmail(string email)
        {
            if (string.IsNullOrEmpty(email)) return email;
            int atIndex = email.IndexOf('@');
            if (atIndex < 0) return email;

            string localPart = email.Substring(0, atIndex);
            string domain = email.Substring(atIndex); // includes '@'

            int visibleChars = Math.Min(3, localPart.Length);
            return localPart.Substring(0, visibleChars) + "***" + domain;
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var sb = new StringBuilder(s.Length);
            foreach (var c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ')
                            sb.AppendFormat("\\u{0:x4}", (int)c);
                        else
                            sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
