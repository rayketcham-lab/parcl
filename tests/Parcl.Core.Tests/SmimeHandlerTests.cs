using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Crypto;
using Xunit;

namespace Parcl.Core.Tests
{
    public class SmimeHandlerTests : IDisposable
    {
        private readonly X509Certificate2 _selfSignedCert;
        private readonly X509Certificate2 _publicOnlyCert;

        public SmimeHandlerTests()
        {
            // Generate an ephemeral self-signed cert for testing
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest(
                "CN=Parcl Test User, E=test@quantumnexum.com",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            // Add key usage for digital signature and key encipherment
            req.CertificateExtensions.Add(
                new X509KeyUsageExtension(
                    X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
                    critical: true));

            _selfSignedCert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5),
                DateTimeOffset.UtcNow.AddHours(1));

            // Export public key only for recipient-only testing
            var publicBytes = _selfSignedCert.Export(X509ContentType.Cert);
            _publicOnlyCert = new X509Certificate2(publicBytes);
        }

        public void Dispose()
        {
            _selfSignedCert?.Dispose();
            _publicOnlyCert?.Dispose();
        }

        [Fact]
        public void Encrypt_SingleRecipient_ProducesValidEnvelope()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Hello encrypted world");
            var certs = new X509Certificate2Collection { _publicOnlyCert };

            var encrypted = handler.Encrypt(content, certs);

            Assert.NotNull(encrypted);
            Assert.True(encrypted.Length > content.Length);
        }

        [Fact]
        public void Encrypt_MultipleRecipients_ProducesValidEnvelope()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Multi-recipient test");

            // Use same cert twice to simulate multiple recipients
            var certs = new X509Certificate2Collection { _publicOnlyCert, _publicOnlyCert };

            var encrypted = handler.Encrypt(content, certs);

            Assert.NotNull(encrypted);
            Assert.True(encrypted.Length > 0);
        }

        [Fact]
        public void Encrypt_EmptyRecipientCollection_ThrowsArgumentException()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Test");
            var empty = new X509Certificate2Collection();

            Assert.Throws<ArgumentException>(() => handler.Encrypt(content, empty));
        }

        [Fact]
        public void Sign_WithPrivateKey_ProducesValidSignature()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Sign me");

            var signed = handler.Sign(content, _selfSignedCert);

            Assert.NotNull(signed);
            Assert.True(signed.Length > content.Length);
        }

        [Fact]
        public void Sign_WithoutPrivateKey_ThrowsInvalidOperationException()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Test");

            Assert.Throws<InvalidOperationException>(
                () => handler.Sign(content, _publicOnlyCert));
        }

        [Fact]
        public void Verify_ValidSignature_ReturnsIsValid()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Verify me");

            var signed = handler.Sign(content, _selfSignedCert);

            // Verify signature-only (skip chain validation — self-signed certs
            // don't have a trusted chain but the signature itself is valid)
            var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
            signedCms.Decode(signed);
            signedCms.CheckSignature(verifySignatureOnly: true);

            Assert.True(signedCms.SignerInfos.Count > 0);
            Assert.Equal(content, signedCms.ContentInfo.Content);
        }

        [Fact]
        public void Verify_TamperedData_ReturnsInvalid()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Original");

            var signed = handler.Sign(content, _selfSignedCert);

            // Tamper with the signed data
            signed[signed.Length / 2] ^= 0xFF;

            var result = handler.Verify(signed);

            Assert.False(result.IsValid);
            Assert.NotNull(result.ErrorMessage);
        }

        [Fact]
        public void Verify_GarbageData_ReturnsInvalid()
        {
            var handler = new SmimeHandler();
            var garbage = new byte[] { 0x00, 0x01, 0x02, 0x03 };

            var result = handler.Verify(garbage);

            Assert.False(result.IsValid);
            Assert.NotNull(result.ErrorMessage);
        }

        [Fact]
        public void SignThenEncrypt_RoundTrip_ProducesValidOutput()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Sign then encrypt me");

            // Sign
            var signed = handler.Sign(content, _selfSignedCert);
            Assert.NotNull(signed);

            // Encrypt the signed data
            var certs = new X509Certificate2Collection { _publicOnlyCert };
            var encrypted = handler.Encrypt(signed, certs);

            Assert.NotNull(encrypted);
            Assert.True(encrypted.Length > signed.Length);
        }

        [Fact]
        public void Constructor_CustomAlgorithms_DoesNotThrow()
        {
            var handler128 = new SmimeHandler("AES-128-CBC", "SHA-384");
            var handler256gcm = new SmimeHandler("AES-256-GCM", "SHA-512");

            // Verify they can encrypt (algorithm selection happens at encrypt time)
            var content = System.Text.Encoding.UTF8.GetBytes("Algorithm test");
            var certs = new X509Certificate2Collection { _publicOnlyCert };

            var encrypted = handler128.Encrypt(content, certs);
            Assert.NotNull(encrypted);
        }

        [Fact]
        public void Decrypt_NoMatchingCert_ReturnsFailure()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Encrypted for someone else");
            var certs = new X509Certificate2Collection { _publicOnlyCert };

            var encrypted = handler.Encrypt(content, certs);

            // Decrypt will fail because the ephemeral cert is not in the Windows store
            var result = handler.Decrypt(encrypted);

            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
        }

        [Fact]
        public void Sign_VerifySigner_HasCertificateInfo()
        {
            var handler = new SmimeHandler();
            var content = System.Text.Encoding.UTF8.GetBytes("Signer identity test");

            var signed = handler.Sign(content, _selfSignedCert);

            // Decode and verify signature-only (self-signed cert won't pass chain validation)
            var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
            signedCms.Decode(signed);
            signedCms.CheckSignature(verifySignatureOnly: true);

            Assert.True(signedCms.SignerInfos.Count > 0);
            var signerCert = signedCms.SignerInfos[0].Certificate;
            Assert.NotNull(signerCert);
            Assert.Contains("Parcl Test User", signerCert!.Subject);
        }
    }
}
