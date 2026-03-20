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
    public class SmimeEndToEndTests : IDisposable
    {
        private readonly X509Certificate2 _signingCert;
        private readonly X509Certificate2 _encryptionCert;
        private readonly SmimeHandler _handler;

        public SmimeEndToEndTests()
        {
            _handler = new SmimeHandler();
            _signingCert = CreateSelfSignedCert("CN=Parcl Test Signer", X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation);
            _encryptionCert = CreateSelfSignedCert("CN=Parcl Test Encryptor", X509KeyUsageFlags.KeyEncipherment | X509KeyUsageFlags.DataEncipherment);
        }

        [Fact]
        public void SignAndVerify_RoundTrip_Succeeds()
        {
            var plaintext = Encoding.UTF8.GetBytes("Confidential message from Parcl.");

            var signed = _handler.Sign(plaintext, _signingCert);
            Assert.NotNull(signed);
            Assert.True(signed.Length > plaintext.Length);

            // Verify signature only (skip chain validation for self-signed test certs)
            var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
            signedCms.Decode(signed);
            signedCms.CheckSignature(verifySignatureOnly: true);

            Assert.Equal(plaintext, signedCms.ContentInfo.Content);
            Assert.NotNull(signedCms.SignerInfos[0].Certificate);
            Assert.Contains("Parcl Test Signer", signedCms.SignerInfos[0].Certificate.Subject);
        }

        [Fact]
        public void Sign_WithoutPrivateKey_Throws()
        {
            var pubOnly = new X509Certificate2(_signingCert.Export(X509ContentType.Cert));
            var data = Encoding.UTF8.GetBytes("test");

            Assert.Throws<InvalidOperationException>(() => _handler.Sign(data, pubOnly));
        }

        [Fact]
        public void Verify_TamperedData_Fails()
        {
            var plaintext = Encoding.UTF8.GetBytes("Original message.");
            var signed = _handler.Sign(plaintext, _signingCert);

            // Tamper with the signed data
            signed[signed.Length / 2] ^= 0xFF;

            var result = _handler.Verify(signed);
            Assert.False(result.IsValid);
            Assert.NotNull(result.ErrorMessage);
        }

        [Fact]
        public void EncryptAndDecrypt_RoundTrip_Succeeds()
        {
            // Import the encryption cert (with private key) into CurrentUser\My so Decrypt can find it
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(_encryptionCert);
            }

            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Secret message for encryption test.");
                var recipientCerts = new X509Certificate2Collection { _encryptionCert };

                var encrypted = _handler.Encrypt(plaintext, recipientCerts);
                Assert.NotNull(encrypted);
                Assert.True(encrypted.Length > plaintext.Length);

                // Should not contain plaintext
                var plaintextStr = Encoding.UTF8.GetString(plaintext);
                var encryptedStr = Encoding.UTF8.GetString(encrypted);
                Assert.DoesNotContain(plaintextStr, encryptedStr);

                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
            }
            finally
            {
                // Clean up: remove the test cert from the store
                using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Remove(_encryptionCert);
                }
            }
        }

        [Fact]
        public void Encrypt_NoRecipients_Throws()
        {
            var data = Encoding.UTF8.GetBytes("test");
            var empty = new X509Certificate2Collection();

            Assert.Throws<ArgumentException>(() => _handler.Encrypt(data, empty));
        }

        [Fact]
        public void SignThenEncrypt_FullPipeline_Succeeds()
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(_encryptionCert);
            }

            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Sign-then-encrypt pipeline test.");

                // Sign
                var signed = _handler.Sign(plaintext, _signingCert);

                // Encrypt the signed data
                var recipientCerts = new X509Certificate2Collection { _encryptionCert };
                var encrypted = _handler.Encrypt(signed, recipientCerts);

                // Decrypt
                var decryptResult = _handler.Decrypt(encrypted);
                var decryptedSigned = decryptResult.Content;

                // Verify (skip chain validation for self-signed test certs)
                var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
                signedCms.Decode(decryptedSigned);
                signedCms.CheckSignature(verifySignatureOnly: true);
                Assert.Equal(plaintext, signedCms.ContentInfo.Content);
            }
            finally
            {
                using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadWrite);
                    store.Remove(_encryptionCert);
                }
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

                // On Windows, export/reimport to make the private key persistable
                return new X509Certificate2(
                    cert.Export(X509ContentType.Pfx, "test"),
                    "test",
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
        }

        public void Dispose()
        {
            _signingCert?.Dispose();
            _encryptionCert?.Dispose();
        }
    }
}
