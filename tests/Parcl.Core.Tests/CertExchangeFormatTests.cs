using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Parcl.Core.Crypto;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// E2E tests for certificate exchange — DER and PEM formats, payload round-trips.
    /// </summary>
    public class CertExchangeFormatTests : IDisposable
    {
        private readonly CertificateStore _certStore;
        private readonly CertExchange _exchange;
        private readonly X509Certificate2 _testCert;

        public CertExchangeFormatTests()
        {
            _certStore = new CertificateStore();
            _exchange = new CertExchange(_certStore);
            _testCert = CreateCert("CN=Exchange Format Test, E=format@parcl.test");
        }

        // =====================================================================
        // DER (.cer) Format
        // =====================================================================

        [Fact]
        public void ExportDer_ProducesBinaryBytes()
        {
            AddToStore(_testCert);
            try
            {
                var derBytes = _certStore.ExportPublicCertificate(_testCert.Thumbprint);
                Assert.NotNull(derBytes);

                // DER starts with ASN.1 SEQUENCE tag (0x30)
                Assert.Equal(0x30, derBytes[0]);

                // Re-parse as cert
                var reimported = new X509Certificate2(derBytes);
                Assert.Equal(_testCert.Subject, reimported.Subject);
                Assert.Equal(_testCert.Thumbprint, reimported.Thumbprint);
                Assert.False(reimported.HasPrivateKey, "DER export should not include private key");
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        // =====================================================================
        // PEM Format
        // =====================================================================

        [Fact]
        public void FormatAsPem_ValidStructure()
        {
            AddToStore(_testCert);
            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                var pem = _exchange.FormatAsAttachment(payload);

                // Check PEM structure
                Assert.StartsWith("-----BEGIN CERTIFICATE-----", pem);
                Assert.Contains("-----END CERTIFICATE-----", pem);

                // Lines should be max 64 chars (PEM standard)
                var lines = pem.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.StartsWith("-----")) continue;
                    Assert.True(line.Length <= 64, $"PEM line exceeds 64 chars: {line.Length}");
                }
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        [Fact]
        public void PemPayload_CanBeDecodedBackToCert()
        {
            AddToStore(_testCert);
            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                var pem = _exchange.FormatAsAttachment(payload);

                // Extract base64 from PEM
                var b64 = pem
                    .Replace("-----BEGIN CERTIFICATE-----", "")
                    .Replace("-----END CERTIFICATE-----", "")
                    .Replace("\r", "")
                    .Replace("\n", "")
                    .Trim();

                var certBytes = Convert.FromBase64String(b64);
                var reimported = new X509Certificate2(certBytes);
                Assert.Equal(_testCert.Subject, reimported.Subject);
                Assert.Equal(_testCert.Thumbprint, reimported.Thumbprint);
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        // =====================================================================
        // CertExchangePayload Tests
        // =====================================================================

        [Fact]
        public void PrepareExport_PopulatesAllFields()
        {
            AddToStore(_testCert);
            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                Assert.NotNull(payload);
                Assert.False(string.IsNullOrEmpty(payload.CertificateData));
                Assert.Equal(_testCert.Thumbprint, payload.Thumbprint);
                Assert.Equal(_testCert.Subject, payload.SenderName);
                Assert.Equal(_testCert.NotAfter, payload.ExpirationDate);
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        [Fact]
        public void ImportFromPayload_AddsToStore_AndReturnsInfo()
        {
            AddToStore(_testCert);
            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);

                // Remove from store to simulate receiving from someone else
                RemoveFromStore(_testCert);

                var info = _exchange.ImportFromPayload(payload);
                Assert.Equal(_testCert.Subject, info.Subject);
                Assert.Equal(_testCert.Thumbprint, info.Thumbprint);

                // Verify it's now in the store
                var found = _certStore.FindByThumbprint(_testCert.Thumbprint);
                Assert.NotNull(found);
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        [Fact]
        public void ImportFromPayload_PublicKeyOnly_NoPrivateKey()
        {
            AddToStore(_testCert);
            try
            {
                var payload = _exchange.PrepareExport(_testCert.Thumbprint);
                RemoveFromStore(_testCert);

                var info = _exchange.ImportFromPayload(payload);
                Assert.False(info.HasPrivateKey, "Imported cert from exchange should not have private key");
            }
            finally
            {
                RemoveFromStore(_testCert);
            }
        }

        [Fact]
        public void FullExchangeWorkflow_ExportImportAndEncrypt()
        {
            // Simulates: Alice exports her cert → Bob imports it → Bob encrypts to Alice → Alice decrypts
            var aliceCert = CreateCert("CN=Alice, E=alice@parcl.test");
            AddToStore(aliceCert);
            try
            {
                // Alice exports her public cert
                var payload = _exchange.PrepareExport(aliceCert.Thumbprint);

                // Bob imports Alice's public cert and uses it to encrypt
                var alicePubBytes = Convert.FromBase64String(payload.CertificateData);
                var alicePub = new X509Certificate2(alicePubBytes);

                var handler = new SmimeHandler();
                var plaintext = Encoding.UTF8.GetBytes("Hello Alice, this is Bob!");
                var encrypted = handler.Encrypt(plaintext, new X509Certificate2Collection { alicePub });

                // Alice decrypts with her private key (in store)
                var decryptResult = handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
                Assert.Equal("Hello Alice, this is Bob!", Encoding.UTF8.GetString(decrypted));
            }
            finally
            {
                RemoveFromStore(aliceCert);
                aliceCert.Dispose();
            }
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private static void AddToStore(X509Certificate2 cert)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(cert);
            }
        }

        private static void RemoveFromStore(X509Certificate2 cert)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var matches = store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, false);
                foreach (X509Certificate2 c in matches)
                    store.Remove(c);
            }
        }

        private static X509Certificate2 CreateCert(string subject)
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(new X509KeyUsageExtension(
                    X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment, critical: true));
                var cert = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));
                return new X509Certificate2(cert.Export(X509ContentType.Pfx, "test"), "test",
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
