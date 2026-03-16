using System;
using System.IO;
using Parcl.Core.Config;
using Xunit;

namespace Parcl.Core.Tests
{
    public class ParclLoggerTests : IDisposable
    {
        private readonly string _logDir;

        public ParclLoggerTests()
        {
            _logDir = Path.Combine(Path.GetTempPath(), "parcl-test-logs-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_logDir);
        }

        [Fact]
        public void Logger_CreatesLogFile()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                logPath = logger.GetLogFilePath();
            }
            Assert.True(File.Exists(logPath), $"Log file should exist at {logPath}");
            Assert.EndsWith(".jsonl", logPath);
        }

        [Fact]
        public void Logger_WritesJsonlFormat()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                logger.Debug("Test", "debug message");
                logger.Info("Test", "info message");
                logger.Warn("Test", "warn message");
                logger.Error("Test", "error message");
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);

            // Each line is a JSON object with structured fields
            Assert.Contains("\"lvl\":\"DEBUG\"", content);
            Assert.Contains("\"lvl\":\"INFO\"", content);
            Assert.Contains("\"lvl\":\"WARN\"", content);
            Assert.Contains("\"lvl\":\"ERROR\"", content);
            Assert.Contains("\"msg\":\"debug message\"", content);
            Assert.Contains("\"msg\":\"info message\"", content);
            Assert.Contains("\"cmp\":\"Test\"", content);

            // Has session ID and PID
            Assert.Contains("\"sid\":\"", content);
            Assert.Contains("\"pid\":", content);
        }

        [Fact]
        public void Logger_IncludesComponentName()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                logger.Info("LDAP", "test lookup");
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);
            Assert.Contains("\"cmp\":\"LDAP\"", content);
        }

        [Fact]
        public void Logger_ErrorWithException_IncludesStructuredError()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                try
                {
                    throw new InvalidOperationException("test failure");
                }
                catch (Exception ex)
                {
                    logger.Error("Test", "operation failed", ex);
                }
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);
            Assert.Contains("\"err\":{", content);
            Assert.Contains("\"type\":\"InvalidOperationException\"", content);
            Assert.Contains("\"msg\":\"test failure\"", content);
            Assert.Contains("\"at\":\"", content);
        }

        [Fact]
        public void Logger_RespectsMinLevel()
        {
            string logPath;
            using (var logger = new ParclLogger(LogLevel.Warn, logDirectory: _logDir))
            {
                logger.Debug("Test", "filtered-debug-should-not-appear");
                logger.Info("Test", "filtered-info-should-not-appear");
                logger.Warn("Test", "this-warn-should-appear");
                logger.Error("Test", "this-error-should-appear");
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);
            Assert.DoesNotContain("filtered-debug-should-not-appear", content);
            Assert.DoesNotContain("filtered-info-should-not-appear", content);
            Assert.Contains("this-warn-should-appear", content);
            Assert.Contains("this-error-should-appear", content);
        }

        [Fact]
        public void Logger_GetLogFilePath_ContainsDate()
        {
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                var path = logger.GetLogFilePath();
                Assert.Contains(DateTime.Now.ToString("yyyy-MM-dd"), path);
                Assert.EndsWith(".jsonl", path);
            }
        }

        [Fact]
        public void Logger_SessionId_IsConsistentAcrossEntries()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                logger.Info("A", "first");
                logger.Info("B", "second");
                logPath = logger.GetLogFilePath();
            }
            var lines = File.ReadAllLines(logPath);
            // All lines from the same session should share the same sid
            // Extract sid from first and last data lines
            Assert.True(lines.Length >= 3, "Should have at least 3 log lines");
            var sid1 = ExtractField(lines[0], "sid");
            var sid2 = ExtractField(lines[1], "sid");
            Assert.Equal(sid1, sid2);
            Assert.Equal(6, sid1.Length); // 3 bytes = 6 hex chars
        }

        [Fact]
        public void Logger_EscapesSpecialCharacters()
        {
            string logPath;
            using (var logger = new ParclLogger(logDirectory: _logDir))
            {
                logger.Info("Test", "quotes \"here\" and\\backslash\nnewline");
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);
            Assert.Contains("quotes \\\"here\\\"", content);
            Assert.Contains("and\\\\backslash", content);
            Assert.Contains("\\n", content);
        }

        private static string ExtractField(string json, string field)
        {
            var key = $"\"{field}\":\"";
            var start = json.IndexOf(key, StringComparison.Ordinal);
            if (start < 0) return string.Empty;
            start += key.Length;
            var end = json.IndexOf('"', start);
            return json.Substring(start, end - start);
        }

        public void Dispose()
        {
            try { Directory.Delete(_logDir, recursive: true); } catch { }
        }
    }
}
