using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Ldap;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    public class CertificateCacheEndToEndTests : IDisposable
    {
        private readonly string _originalCacheFile;
        private readonly string _backupCacheFile;
        private readonly string _cacheDir;

        public CertificateCacheEndToEndTests()
        {
            _cacheDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl");
            _originalCacheFile = Path.Combine(_cacheDir, "cert-cache.json");
            _backupCacheFile = Path.Combine(_cacheDir, "cert-cache.json.testbackup");

            // Back up existing cache if present
            if (File.Exists(_originalCacheFile))
                File.Copy(_originalCacheFile, _backupCacheFile, overwrite: true);
        }

        [Fact]
        public void AddAndGet_RoundTrip()
        {
            var cache = new CertificateCache(expirationHours: 1, maxEntries: 100);
            var certs = CreateTestCerts("test@example.com");

            cache.Add("test@example.com", certs);
            var retrieved = cache.Get("test@example.com");

            Assert.NotNull(retrieved);
            Assert.Equal(certs.Count, retrieved.Count);
            Assert.Equal(certs[0].Thumbprint, retrieved[0].Thumbprint);
        }

        [Fact]
        public void Get_CaseInsensitive()
        {
            var cache = new CertificateCache(expirationHours: 1, maxEntries: 100);
            cache.Add("Test@Example.COM", CreateTestCerts("Test@Example.COM"));

            var result = cache.Get("test@example.com");
            Assert.NotNull(result);
        }

        [Fact]
        public void Get_Expired_ReturnsNull()
        {
            // Use 0 expiration hours — entries expire immediately
            var cache = new CertificateCache(expirationHours: 0, maxEntries: 100);
            cache.Add("expired@example.com", CreateTestCerts("expired@example.com"));

            var result = cache.Get("expired@example.com");
            Assert.Null(result);
        }

        [Fact]
        public void Remove_DeletesEntry()
        {
            var cache = new CertificateCache(expirationHours: 1, maxEntries: 100);
            cache.Add("remove@example.com", CreateTestCerts("remove@example.com"));

            cache.Remove("remove@example.com");
            var result = cache.Get("remove@example.com");
            Assert.Null(result);
        }

        [Fact]
        public void Clear_RemovesAllEntries()
        {
            var cache = new CertificateCache(expirationHours: 1, maxEntries: 100);
            cache.Add("a@example.com", CreateTestCerts("a@example.com"));
            cache.Add("b@example.com", CreateTestCerts("b@example.com"));

            cache.Clear();

            Assert.Null(cache.Get("a@example.com"));
            Assert.Null(cache.Get("b@example.com"));
        }

        [Fact]
        public void Eviction_RemovesOldestWhenOverMax()
        {
            var cache = new CertificateCache(expirationHours: 24, maxEntries: 2);
            cache.Add("first@example.com", CreateTestCerts("first@example.com"));
            cache.Add("second@example.com", CreateTestCerts("second@example.com"));
            cache.Add("third@example.com", CreateTestCerts("third@example.com"));

            // "first" should have been evicted
            Assert.Null(cache.Get("first@example.com"));
            Assert.NotNull(cache.Get("third@example.com"));
        }

        [Fact]
        public void Persistence_SurvivesReload()
        {
            var cache1 = new CertificateCache(expirationHours: 24, maxEntries: 100);
            cache1.Add("persist@example.com", CreateTestCerts("persist@example.com"));

            // Create a new cache instance — it should load from disk
            var cache2 = new CertificateCache(expirationHours: 24, maxEntries: 100);
            var result = cache2.Get("persist@example.com");

            Assert.NotNull(result);
            Assert.Single(result);

            // Clean up
            cache2.Clear();
        }

        private List<CertificateInfo> CreateTestCerts(string email)
        {
            return new List<CertificateInfo>
            {
                new CertificateInfo
                {
                    Thumbprint = (Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")).Substring(0, 40).ToUpperInvariant(),
                    Subject = $"CN={email}",
                    Issuer = "CN=Test CA",
                    Email = email,
                    NotBefore = DateTime.UtcNow.AddDays(-1),
                    NotAfter = DateTime.UtcNow.AddYears(1),
                    SerialNumber = "01",
                    KeyUsage = X509KeyUsageFlags.KeyEncipherment,
                    HasPrivateKey = false
                }
            };
        }

        public void Dispose()
        {
            // Restore original cache
            if (File.Exists(_backupCacheFile))
            {
                File.Copy(_backupCacheFile, _originalCacheFile, overwrite: true);
                File.Delete(_backupCacheFile);
            }
        }
    }
}
