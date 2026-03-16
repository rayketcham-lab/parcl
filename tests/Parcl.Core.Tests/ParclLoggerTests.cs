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
            _logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Parcl", "logs");
        }

        [Fact]
        public void Logger_CreatesLogFile()
        {
            string logPath;
            using (var logger = new ParclLogger())
            {
                logPath = logger.GetLogFilePath();
            }
            Assert.True(File.Exists(logPath), $"Log file should exist at {logPath}");
        }

        [Fact]
        public void Logger_WritesAllLevels()
        {
            string logPath;
            using (var logger = new ParclLogger())
            {
                logger.Debug("Test", "debug message");
                logger.Info("Test", "info message");
                logger.Warn("Test", "warn message");
                logger.Error("Test", "error message");
                logPath = logger.GetLogFilePath();
            }
            // Dispose flushes all buffered writes, safe to read now
            var content = File.ReadAllText(logPath);

            Assert.Contains("[DEBUG]", content);
            Assert.Contains("[INFO ]", content);
            Assert.Contains("[WARN ]", content);
            Assert.Contains("[ERROR]", content);
            Assert.Contains("debug message", content);
            Assert.Contains("info message", content);
        }

        [Fact]
        public void Logger_IncludesComponentName()
        {
            string logPath;
            using (var logger = new ParclLogger())
            {
                logger.Info("LDAP", "test lookup");
                logPath = logger.GetLogFilePath();
            }
            var content = File.ReadAllText(logPath);
            Assert.Contains("[LDAP", content);
        }

        [Fact]
        public void Logger_ErrorWithException_IncludesExceptionType()
        {
            string logPath;
            using (var logger = new ParclLogger())
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
            Assert.Contains("InvalidOperationException", content);
            Assert.Contains("test failure", content);
        }

        [Fact]
        public void Logger_RespectsMinLevel()
        {
            // Use a unique temp log to avoid cross-test contamination
            string logPath;
            using (var logger = new ParclLogger(LogLevel.Warn))
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
            using (var logger = new ParclLogger())
            {
                var path = logger.GetLogFilePath();
                Assert.Contains(DateTime.Now.ToString("yyyy-MM-dd"), path);
                Assert.EndsWith(".log", path);
            }
        }

        public void Dispose() { }
    }
}
