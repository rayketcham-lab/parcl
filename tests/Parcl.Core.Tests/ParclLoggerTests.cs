using System;
using System.IO;
using System.Threading;
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
            using (var logger = new ParclLogger())
            {
                var logPath = logger.GetLogFilePath();
                Assert.True(File.Exists(logPath), $"Log file should exist at {logPath}");
            }
        }

        [Fact]
        public void Logger_WritesAllLevels()
        {
            using (var logger = new ParclLogger())
            {
                logger.Debug("Test", "debug message");
                logger.Info("Test", "info message");
                logger.Warn("Test", "warn message");
                logger.Error("Test", "error message");
            }

            // Logger flushes on dispose, read the file
            var logFile = Path.Combine(_logDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");
            var content = File.ReadAllText(logFile);

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
            using (var logger = new ParclLogger())
            {
                logger.Info("LDAP", "test lookup");
            }

            var logFile = Path.Combine(_logDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");
            var content = File.ReadAllText(logFile);
            Assert.Contains("[LDAP", content);
        }

        [Fact]
        public void Logger_ErrorWithException_IncludesExceptionType()
        {
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
            }

            var logFile = Path.Combine(_logDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");
            var content = File.ReadAllText(logFile);
            Assert.Contains("InvalidOperationException", content);
            Assert.Contains("test failure", content);
        }

        [Fact]
        public void Logger_RespectsMinLevel()
        {
            var logFile = Path.Combine(_logDir, $"parcl-{DateTime.Now:yyyy-MM-dd}.log");

            // Delete existing to start clean
            if (File.Exists(logFile)) File.Delete(logFile);

            using (var logger = new ParclLogger(LogLevel.Warn))
            {
                logger.Debug("Test", "should not appear");
                logger.Info("Test", "should not appear either");
                logger.Warn("Test", "this should appear");
                logger.Error("Test", "this too");
            }

            var content = File.ReadAllText(logFile);
            Assert.DoesNotContain("should not appear", content);
            Assert.Contains("this should appear", content);
            Assert.Contains("this too", content);
        }

        public void Dispose()
        {
            // Clean up test logs
        }
    }
}
