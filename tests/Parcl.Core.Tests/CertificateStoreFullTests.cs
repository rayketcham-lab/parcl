using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Crypto;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// E2E tests for CertificateStore and CertificateInfo — filtering, import/export, key usage classification.
    /// </summary>
    [Collection("CertStore")]
    public class CertificateStoreFullTests : IDisposable
    {
        private readonly CertificateStore _store;
        private readonly List<X509Certificate2> _testCerts = new List<X509Certificate2>();

        public CertificateStoreFullTests()
        {
            _store = new CertificateStore();
        }

        // =====================================================================
        // CertificateInfo Key Usage Classification (does not require cert store)
        // =====================================================================

        [Fact]
        public void CertificateInfo_SigningOnly_ClassifiedCorrectly()
        {
            var cert = CreateCert("CN=Signing Only", X509KeyUsageFlags.DigitalSignature);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsSigningCert);
            Assert.False(info.IsEncryptionCert);
        }

        [Fact]
        public void CertificateInfo_NonRepudiation_ClassifiedAsSigning()
        {
            var cert = CreateCert("CN=NonRepudiation", X509KeyUsageFlags.NonRepudiation);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsSigningCert);
            Assert.False(info.IsEncryptionCert);
        }

        [Fact]
        public void CertificateInfo_KeyEncipherment_ClassifiedAsEncryption()
        {
            var cert = CreateCert("CN=KeyEncipherment", X509KeyUsageFlags.KeyEncipherment);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsEncryptionCert);
            Assert.False(info.IsSigningCert);
        }

        [Fact]
        public void CertificateInfo_DataEncipherment_ClassifiedAsEncryption()
        {
            var cert = CreateCert("CN=DataEncipherment", X509KeyUsageFlags.DataEncipherment);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsEncryptionCert);
            Assert.False(info.IsSigningCert);
        }

        [Fact]
        public void CertificateInfo_DualUse_ClassifiedAsBoth()
        {
            var cert = CreateCert("CN=DualUse",
                X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsSigningCert);
            Assert.True(info.IsEncryptionCert);
        }

        [Fact]
        public void CertificateInfo_AllKeyUsageFlags_Correct()
        {
            var cert = CreateCert("CN=AllFlags",
                X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation |
                X509KeyUsageFlags.KeyEncipherment | X509KeyUsageFlags.DataEncipherment);
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsSigningCert);
            Assert.True(info.IsEncryptionCert);
        }

        // =====================================================================
        // CertificateInfo Validity States
        // =====================================================================

        [Fact]
        public void CertificateInfo_ValidCert_IsValid()
        {
            // Use wide validity window to avoid UTC/local timezone edge cases
            var cert = CreateCert("CN=Valid", X509KeyUsageFlags.DigitalSignature,
                notBefore: DateTimeOffset.UtcNow.AddDays(-1),
                notAfter: DateTimeOffset.UtcNow.AddDays(365));
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsValid);
            Assert.False(info.IsExpired);
            Assert.False(info.IsNotYetValid);
        }

        [Fact]
        public void CertificateInfo_ExpiredCert_IsExpired()
        {
            var cert = CreateCert("CN=Expired", X509KeyUsageFlags.DigitalSignature,
                notBefore: DateTimeOffset.UtcNow.AddDays(-30),
                notAfter: DateTimeOffset.UtcNow.AddDays(-1));
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsExpired);
            Assert.False(info.IsValid);
        }

        [Fact]
        public void CertificateInfo_NotYetValidCert_IsNotYetValid()
        {
            var cert = CreateCert("CN=Future", X509KeyUsageFlags.DigitalSignature,
                notBefore: DateTimeOffset.UtcNow.AddDays(1),
                notAfter: DateTimeOffset.UtcNow.AddDays(365));
            _testCerts.Add(cert);
            var info = CertificateInfo.FromX509(cert);

            Assert.True(info.IsNotYetValid);
            Assert.False(info.IsValid);
        }

        // =====================================================================
        // FromX509 Field Mapping
        // =====================================================================

        [Fact]
        public void FromX509_MapsAllFields()
        {
            var cert = CreateCert("CN=Mapping Test",
                X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment);
            _testCerts.Add(cert);

            var info = CertificateInfo.FromX509(cert);
            Assert.Equal(cert.Thumbprint, info.Thumbprint);
            Assert.Equal(cert.Subject, info.Subject);
            Assert.Equal(cert.Issuer, info.Issuer);
            Assert.Equal(cert.NotBefore, info.NotBefore);
            Assert.Equal(cert.NotAfter, info.NotAfter);
            Assert.Equal(cert.SerialNumber, info.SerialNumber);
            Assert.Equal(cert.HasPrivateKey, info.HasPrivateKey);
            Assert.NotNull(info.RawData);
            Assert.True(info.IsSigningCert);
            Assert.True(info.IsEncryptionCert);
        }

        [Fact]
        public void FromX509_ToString_Format()
        {
            var cert = CreateCert("CN=ToString Test", X509KeyUsageFlags.DigitalSignature);
            _testCerts.Add(cert);

            var info = CertificateInfo.FromX509(cert);
            var str = info.ToString();

            Assert.Contains("CN=ToString Test", str);
            Assert.Contains(cert.Thumbprint.Substring(0, 8), str);
            Assert.Contains(cert.NotAfter.ToString("yyyy-MM-dd"), str);
        }

        // =====================================================================
        // CertificateStore — FindByThumbprint
        // =====================================================================

        [Fact]
        public void FindByThumbprint_ExactMatch()
        {
            var cert = CreateAndAdd("CN=Thumbprint Find Test",
                X509KeyUsageFlags.DigitalSignature);
            try
            {
                var found = _store.FindByThumbprint(cert.Thumbprint);
                Assert.NotNull(found);
                Assert.Equal(cert.Subject, found.Subject);
            }
            finally
            {
                Cleanup();
            }
        }

        [Fact]
        public void FindByThumbprint_NotFound_ReturnsNull()
        {
            var result = _store.FindByThumbprint("DEADBEEFDEADBEEFDEADBEEFDEADBEEFDEADBEEF");
            Assert.Null(result);
        }

        // =====================================================================
        // FindByEmail
        // =====================================================================

        [Fact]
        public void FindByEmail_ReturnsLatestValid()
        {
            var cert1 = CreateAndAdd("CN=emailtest@parcl.test",
                X509KeyUsageFlags.KeyEncipherment,
                notAfter: DateTimeOffset.UtcNow.AddDays(30));
            var cert2 = CreateAndAdd("CN=emailtest@parcl.test",
                X509KeyUsageFlags.KeyEncipherment,
                notAfter: DateTimeOffset.UtcNow.AddDays(365));

            try
            {
                var found = _store.FindByEmail("emailtest@parcl.test");
                Assert.NotNull(found);
                Assert.Equal(cert2.Thumbprint, found.Thumbprint);
            }
            finally
            {
                Cleanup();
            }
        }

        [Fact]
        public void FindByEmail_NotFound_ReturnsNull()
        {
            var result = _store.FindByEmail("nonexistent@parcl.test");
            Assert.Null(result);
        }

        // =====================================================================
        // Import / Export
        // =====================================================================

        [Fact]
        public void ImportCertificate_ThenExport_RoundTrip()
        {
            var cert = CreateCert("CN=Import Export Test",
                X509KeyUsageFlags.KeyEncipherment);
            _testCerts.Add(cert);

            var pubBytes = cert.Export(X509ContentType.Cert);
            _store.ImportCertificate(pubBytes);

            try
            {
                var exported = _store.ExportPublicCertificate(cert.Thumbprint);
                Assert.NotNull(exported);

                var reimported = new X509Certificate2(exported);
                Assert.Equal(cert.Thumbprint, reimported.Thumbprint);
                Assert.False(reimported.HasPrivateKey);
            }
            finally
            {
                Cleanup();
            }
        }

        [Fact]
        public void ExportPublicCertificate_NonExistent_ReturnsNull()
        {
            var result = _store.ExportPublicCertificate("0000000000000000000000000000000000000000");
            Assert.Null(result);
        }

        // =====================================================================
        // Store listing
        // =====================================================================

        [Fact]
        public void GetAllCertificates_DoesNotThrow()
        {
            var all = _store.GetAllCertificates();
            Assert.NotNull(all);
        }

        [Fact]
        public void GetSigningCertificates_DoesNotThrow()
        {
            var signing = _store.GetSigningCertificates();
            Assert.NotNull(signing);
        }

        [Fact]
        public void GetEncryptionCertificates_DoesNotThrow()
        {
            var encryption = _store.GetEncryptionCertificates();
            Assert.NotNull(encryption);
        }

        [Fact]
        public void GetAllCertificates_OrderedByExpirationDescending()
        {
            var all = _store.GetAllCertificates();
            if (all.Count >= 2)
            {
                for (int i = 0; i < all.Count - 1; i++)
                {
                    Assert.True(all[i].NotAfter >= all[i + 1].NotAfter,
                        "Certs should be ordered by NotAfter descending");
                }
            }
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private X509Certificate2 CreateAndAdd(string subject, X509KeyUsageFlags keyUsage,
            DateTimeOffset? notBefore = null, DateTimeOffset? notAfter = null)
        {
            var cert = CreateCert(subject, keyUsage, notBefore, notAfter);
            _testCerts.Add(cert);
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(cert);
            }
            return cert;
        }

        private static X509Certificate2 CreateCert(string subject, X509KeyUsageFlags keyUsage,
            DateTimeOffset? notBefore = null, DateTimeOffset? notAfter = null)
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(new X509KeyUsageExtension(keyUsage, critical: true));
                var cert = request.CreateSelfSigned(
                    notBefore ?? DateTimeOffset.UtcNow.AddMinutes(-5),
                    notAfter ?? DateTimeOffset.UtcNow.AddHours(1));
                return new X509Certificate2(cert.Export(X509ContentType.Pfx, "test"), "test",
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
        }

        private void Cleanup()
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                foreach (var cert in _testCerts)
                {
                    var matches = store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, false);
                    foreach (X509Certificate2 c in matches)
                        store.Remove(c);
                }
            }
        }

        public void Dispose()
        {
            Cleanup();
            foreach (var cert in _testCerts) cert?.Dispose();
            _store?.Dispose();
        }
    }
}
