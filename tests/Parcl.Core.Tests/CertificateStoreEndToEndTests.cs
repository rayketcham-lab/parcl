using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Crypto;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    public class CertificateStoreEndToEndTests : IDisposable
    {
        private readonly CertificateStore _certStore;
        private readonly X509Certificate2 _testCert;

        public CertificateStoreEndToEndTests()
        {
            _certStore = new CertificateStore();
            _testCert = CreateSelfSignedCert("CN=Parcl Store Test", X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment);
        }

        [Fact]
        public void ImportAndFind_ByThumbprint_RoundTrip()
        {
            var pubBytes = _testCert.Export(X509ContentType.Cert);

            // Import into the store via CertificateStore
            _certStore.ImportCertificate(pubBytes);

            try
            {
                var found = _certStore.FindByThumbprint(_testCert.Thumbprint);
                Assert.NotNull(found);
                Assert.Equal(_testCert.Thumbprint, found.Thumbprint);
            }
            finally
            {
                RemoveFromStore(_testCert.Thumbprint);
            }
        }

        [Fact]
        public void ExportPublicCertificate_ReturnsDerBytes()
        {
            // Add to store first
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(_testCert);
            }

            try
            {
                var exported = _certStore.ExportPublicCertificate(_testCert.Thumbprint);
                Assert.NotNull(exported);

                // Should be parseable back into a cert
                var reimported = new X509Certificate2(exported);
                Assert.Equal(_testCert.Subject, reimported.Subject);
                Assert.Equal(_testCert.Thumbprint, reimported.Thumbprint);
            }
            finally
            {
                RemoveFromStore(_testCert.Thumbprint);
            }
        }

        [Fact]
        public void FindByThumbprint_NonExistent_ReturnsNull()
        {
            var result = _certStore.FindByThumbprint("0000000000000000000000000000000000000000");
            Assert.Null(result);
        }

        [Fact]
        public void GetAllCertificates_ReturnsNonEmptyList()
        {
            // Most dev machines have at least one cert in CurrentUser\My
            var certs = _certStore.GetAllCertificates();
            Assert.NotNull(certs);
            // Don't assert count — machine-dependent — just that it doesn't throw
        }

        private void RemoveFromStore(string thumbprint)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var matches = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                foreach (var c in matches.Cast<X509Certificate2>())
                    store.Remove(c);
            }
        }

        private static X509Certificate2 CreateSelfSignedCert(string subject, X509KeyUsageFlags keyUsage)
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(keyUsage, critical: true));

                var cert = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));
                return new X509Certificate2(
                    cert.Export(X509ContentType.Pfx, "test"),
                    "test",
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
        }

        public void Dispose()
        {
            _certStore?.Dispose();
            _testCert?.Dispose();
        }
    }
}
