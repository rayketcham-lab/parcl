using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Crypto;
using Xunit;

namespace Parcl.Core.Tests
{
    public class CertExchangeEndToEndTests : IDisposable
    {
        private readonly CertificateStore _certStore;
        private readonly CertExchange _exchange;
        private readonly X509Certificate2 _testCert;

        public CertExchangeEndToEndTests()
        {
            _certStore = new CertificateStore();
            _exchange = new CertExchange(_certStore);
            _testCert = CreateSelfSignedCert("CN=Parcl Exchange Test");
        }

        [Fact]
        public void PrepareExport_And_ImportFromPayload_RoundTrip()
        {
            // Put the cert in the store so PrepareExport can find it
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(_testCert);
            }

            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                Assert.NotNull(payload);
                Assert.False(string.IsNullOrEmpty(payload.CertificateData));
                Assert.Equal(_testCert.Thumbprint, payload.Thumbprint);

                // Remove the cert first so import is meaningful
                RemoveFromStore(_testCert.Thumbprint);

                var imported = _exchange.ImportFromPayload(payload);
                Assert.Equal(_testCert.Subject, imported.Subject);
                Assert.Equal(_testCert.Thumbprint, imported.Thumbprint);
            }
            finally
            {
                RemoveFromStore(_testCert.Thumbprint);
            }
        }

        [Fact]
        public void FormatAsAttachment_ProducesValidPem()
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(_testCert);
            }

            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                var pem = _exchange.FormatAsAttachment(payload);

                Assert.StartsWith("-----BEGIN CERTIFICATE-----", pem);
                Assert.Contains("-----END CERTIFICATE-----", pem);

                // Each line between markers should be <= 64 chars
                var lines = pem.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.StartsWith("-----")) continue;
                    Assert.True(line.Length <= 64, $"PEM line too long: {line.Length} chars");
                }
            }
            finally
            {
                RemoveFromStore(_testCert.Thumbprint);
            }
        }

        [Fact]
        public void PrepareExport_NonExistentThumbprint_Throws()
        {
            Assert.Throws<InvalidOperationException>(
                () => _exchange.PrepareExport("0000000000000000000000000000000000000000"));
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

        private static X509Certificate2 CreateSelfSignedCert(string subject)
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(
                        X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
                        critical: true));

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
