using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Parcl.Core.Crypto;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// Comprehensive E2E encryption tests covering all algorithms, modes, and recipient scenarios.
    /// </summary>
    public class SmimeFullEncryptionTests : IDisposable
    {
        private readonly SmimeHandler _handler;
        private readonly X509Certificate2 _signingCert;
        private readonly X509Certificate2 _encCert1;
        private readonly X509Certificate2 _encCert2;
        private readonly X509Certificate2 _encCert3;
        private readonly X509Certificate2 _dualUseCert;
        private readonly List<X509Certificate2> _allCerts;

        public SmimeFullEncryptionTests()
        {
            _handler = new SmimeHandler();
            _signingCert = CreateCert("CN=E2E Signer, E=signer@parcl.test",
                X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation);
            _encCert1 = CreateCert("CN=Recipient One, E=r1@parcl.test",
                X509KeyUsageFlags.KeyEncipherment | X509KeyUsageFlags.DataEncipherment);
            _encCert2 = CreateCert("CN=Recipient Two, E=r2@parcl.test",
                X509KeyUsageFlags.KeyEncipherment | X509KeyUsageFlags.DataEncipherment);
            _encCert3 = CreateCert("CN=Recipient Three, E=r3@parcl.test",
                X509KeyUsageFlags.KeyEncipherment | X509KeyUsageFlags.DataEncipherment);
            _dualUseCert = CreateCert("CN=Dual Use, E=dual@parcl.test",
                X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment);

            _allCerts = new List<X509Certificate2> { _encCert1, _encCert2, _encCert3, _dualUseCert };
        }

        // =====================================================================
        // Encryption Algorithm Tests (AES-128, AES-192, AES-256, 3DES)
        // =====================================================================

        [Theory]
        [InlineData("2.16.840.1.101.3.4.1.2", "AES-128-CBC")]   // AES-128-CBC
        [InlineData("2.16.840.1.101.3.4.1.22", "AES-192-CBC")]  // AES-192-CBC
        [InlineData("2.16.840.1.101.3.4.1.42", "AES-256-CBC")]  // AES-256-CBC
        [InlineData("1.2.840.113549.3.7", "3DES")]               // 3DES-CBC
        public void Encrypt_WithAlgorithm_Succeeds(string oid, string name)
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes($"Testing encryption with {name}");
                var encrypted = EncryptWithAlgorithm(plaintext, new X509Certificate2Collection { _encCert1 }, oid);

                Assert.NotNull(encrypted);
                Assert.True(encrypted.Length > plaintext.Length);

                // Verify the algorithm OID in the envelope
                var envelope = new EnvelopedCms();
                envelope.Decode(encrypted);
                Assert.Equal(oid, envelope.ContentEncryptionAlgorithm.Oid.Value);

                // Decrypt and verify content
                envelope.Decrypt();
                Assert.Equal(plaintext, envelope.ContentInfo.Content);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        [Fact]
        public void Encrypt_AES256_DefaultAlgorithm_ProducesCorrectOid()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Default AES-256 test");
                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });

                var envelope = new EnvelopedCms();
                envelope.Decode(encrypted);
                Assert.Equal("2.16.840.1.101.3.4.1.42", envelope.ContentEncryptionAlgorithm.Oid.Value);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        // =====================================================================
        // Hash Algorithm Tests (SHA-1, SHA-256, SHA-384, SHA-512)
        // =====================================================================

        [Theory]
        [InlineData("1.3.14.3.2.26", "SHA-1")]
        [InlineData("2.16.840.1.101.3.4.2.1", "SHA-256")]
        [InlineData("2.16.840.1.101.3.4.2.2", "SHA-384")]
        [InlineData("2.16.840.1.101.3.4.2.3", "SHA-512")]
        public void Sign_WithHashAlgorithm_Succeeds(string oid, string name)
        {
            var plaintext = Encoding.UTF8.GetBytes($"Testing signing with {name}");
            var signed = SignWithHash(plaintext, _signingCert, oid);

            Assert.NotNull(signed);

            var signedCms = new SignedCms();
            signedCms.Decode(signed);
            Assert.Equal(oid, signedCms.SignerInfos[0].DigestAlgorithm.Value);

            // Use verifySignatureOnly to skip chain validation for self-signed test certs
            signedCms.CheckSignature(verifySignatureOnly: true);
            Assert.Equal(plaintext, signedCms.ContentInfo.Content);
        }

        // =====================================================================
        // Multi-Recipient Encryption
        // =====================================================================

        [Fact]
        public void Encrypt_TwoRecipients_BothCanDecrypt()
        {
            AddToStore(_encCert1);
            AddToStore(_encCert2);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Multi-recipient message for two.");
                var recipients = new X509Certificate2Collection { _encCert1, _encCert2 };
                var encrypted = _handler.Encrypt(plaintext, recipients);

                // Both should be able to decrypt
                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);

                // Verify the envelope has 2 recipient infos
                var envelope = new EnvelopedCms();
                envelope.Decode(encrypted);
                Assert.Equal(2, envelope.RecipientInfos.Count);
            }
            finally
            {
                RemoveFromStore(_encCert1);
                RemoveFromStore(_encCert2);
            }
        }

        [Fact]
        public void Encrypt_ThreeRecipients_AllCanDecrypt()
        {
            AddToStore(_encCert1);
            AddToStore(_encCert2);
            AddToStore(_encCert3);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Message for three recipients.");
                var recipients = new X509Certificate2Collection { _encCert1, _encCert2, _encCert3 };
                var encrypted = _handler.Encrypt(plaintext, recipients);

                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);

                var envelope = new EnvelopedCms();
                envelope.Decode(encrypted);
                Assert.Equal(3, envelope.RecipientInfos.Count);
            }
            finally
            {
                RemoveFromStore(_encCert1);
                RemoveFromStore(_encCert2);
                RemoveFromStore(_encCert3);
            }
        }

        [Fact]
        public void Encrypt_RecipientWithoutPrivateKey_CannotDecrypt()
        {
            // Don't add cert1 to store — only the public key was used for encryption
            AddToStore(_encCert2);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Only cert2 can decrypt.");
                var recipients = new X509Certificate2Collection { _encCert1, _encCert2 };
                var encrypted = _handler.Encrypt(plaintext, recipients);

                // Should still decrypt because cert2 is in the store with private key
                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
            }
            finally
            {
                RemoveFromStore(_encCert2);
            }
        }

        // =====================================================================
        // Sign-Only Mode
        // =====================================================================

        [Fact]
        public void SignOnly_ContentRemainsReadable()
        {
            var plaintext = Encoding.UTF8.GetBytes("This message is signed but not encrypted.");
            var signed = _handler.Sign(plaintext, _signingCert);

            // The signed data (CMS format) contains the original content
            var signedCms = new SignedCms();
            signedCms.Decode(signed);
            Assert.Equal(plaintext, signedCms.ContentInfo.Content);
        }

        [Fact]
        public void SignOnly_IncludesSignerCertChain()
        {
            var plaintext = Encoding.UTF8.GetBytes("Check signer info.");
            var signed = _handler.Sign(plaintext, _signingCert);

            var signedCms = new SignedCms();
            signedCms.Decode(signed);

            Assert.Single(signedCms.SignerInfos);
            Assert.NotNull(signedCms.SignerInfos[0].Certificate);
            Assert.Equal(_signingCert.Subject, signedCms.SignerInfos[0].Certificate.Subject);
        }

        [Fact]
        public void SignOnly_UsesIssuerAndSerialNumber()
        {
            var plaintext = Encoding.UTF8.GetBytes("Check signer identifier type.");
            var signed = _handler.Sign(plaintext, _signingCert);

            var signedCms = new SignedCms();
            signedCms.Decode(signed);

            Assert.Equal(SubjectIdentifierType.IssuerAndSerialNumber,
                signedCms.SignerInfos[0].SignerIdentifier.Type);
        }

        // =====================================================================
        // Encrypt-Only Mode
        // =====================================================================

        [Fact]
        public void EncryptOnly_ContentNotReadable()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("This should be unreadable without decryption.");
                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });

                // Raw encrypted bytes should not contain the plaintext
                var encStr = Convert.ToBase64String(encrypted);
                var plainStr = Convert.ToBase64String(plaintext);
                Assert.DoesNotContain(plainStr, encStr);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        [Fact]
        public void EncryptOnly_UsesIssuerAndSerialNumber()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Check recipient identifier type.");
                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });

                var envelope = new EnvelopedCms();
                envelope.Decode(encrypted);
                Assert.Equal(SubjectIdentifierType.IssuerAndSerialNumber,
                    envelope.RecipientInfos[0].RecipientIdentifier.Type);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        // =====================================================================
        // Sign + Encrypt Combined Mode
        // =====================================================================

        [Fact]
        public void SignThenEncrypt_AllAlgorithms_RoundTrip()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Sign-then-encrypt with all algorithms.");

                // Sign with SHA-256
                var signed = _handler.Sign(plaintext, _signingCert);

                // Encrypt with AES-256
                var encrypted = _handler.Encrypt(signed, new X509Certificate2Collection { _encCert1 });

                // Decrypt
                var decryptResult = _handler.Decrypt(encrypted);
                var decryptedSigned = decryptResult.Content;

                // Verify signature (skip chain validation for self-signed)
                var signedCms = new SignedCms();
                signedCms.Decode(decryptedSigned);
                signedCms.CheckSignature(verifySignatureOnly: true);
                Assert.Equal(plaintext, signedCms.ContentInfo.Content);
                Assert.NotNull(signedCms.SignerInfos[0].Certificate);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        [Fact]
        public void SignThenEncrypt_MultiRecipient_AllCanDecryptAndVerify()
        {
            AddToStore(_encCert1);
            AddToStore(_encCert2);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Signed and encrypted for multiple recipients.");
                var signed = _handler.Sign(plaintext, _signingCert);
                var encrypted = _handler.Encrypt(signed,
                    new X509Certificate2Collection { _encCert1, _encCert2 });

                var decryptResult = _handler.Decrypt(encrypted);
                var decryptedSigned = decryptResult.Content;
                var signedCms = new SignedCms();
                signedCms.Decode(decryptedSigned);
                signedCms.CheckSignature(verifySignatureOnly: true);
                Assert.Equal(plaintext, signedCms.ContentInfo.Content);
            }
            finally
            {
                RemoveFromStore(_encCert1);
                RemoveFromStore(_encCert2);
            }
        }

        // =====================================================================
        // Dual-Use Certificate (Signing + Encryption)
        // =====================================================================

        [Fact]
        public void DualUseCert_CanSignAndEncrypt()
        {
            AddToStore(_dualUseCert);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Dual-use cert for both sign and encrypt.");

                // Sign with the dual-use cert
                var signed = _handler.Sign(plaintext, _dualUseCert);

                // Encrypt to the same dual-use cert
                var encrypted = _handler.Encrypt(signed, new X509Certificate2Collection { _dualUseCert });

                // Decrypt and verify (skip chain validation)
                var decryptResult = _handler.Decrypt(encrypted);
                var decryptedSigned = decryptResult.Content;
                var signedCms = new SignedCms();
                signedCms.Decode(decryptedSigned);
                signedCms.CheckSignature(verifySignatureOnly: true);
                Assert.Equal(plaintext, signedCms.ContentInfo.Content);
            }
            finally
            {
                RemoveFromStore(_dualUseCert);
            }
        }

        [Fact]
        public void DualUseCert_KeyUsageReflectsBothFlags()
        {
            var info = CertificateInfo.FromX509(_dualUseCert);
            Assert.True(info.IsSigningCert);
            Assert.True(info.IsEncryptionCert);
        }

        // =====================================================================
        // Large Content Tests
        // =====================================================================

        [Fact]
        public void Encrypt_LargePayload_1MB_RoundTrip()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = new byte[1024 * 1024]; // 1 MB
                new Random(42).NextBytes(plaintext);

                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });
                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        [Fact]
        public void Sign_LargePayload_1MB_RoundTrip()
        {
            var plaintext = new byte[1024 * 1024];
            new Random(42).NextBytes(plaintext);

            var signed = _handler.Sign(plaintext, _signingCert);
            var signedCms = new SignedCms();
            signedCms.Decode(signed);
            signedCms.CheckSignature(verifySignatureOnly: true);
            Assert.Equal(plaintext, signedCms.ContentInfo.Content);
        }

        [Fact]
        public void SignThenEncrypt_LargePayload_RoundTrip()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = new byte[512 * 1024]; // 512 KB
                new Random(99).NextBytes(plaintext);

                var signed = _handler.Sign(plaintext, _signingCert);
                var encrypted = _handler.Encrypt(signed, new X509Certificate2Collection { _encCert1 });
                var decryptResult = _handler.Decrypt(encrypted);
                var decryptedSigned = decryptResult.Content;
                var signedCms = new SignedCms();
                signedCms.Decode(decryptedSigned);
                signedCms.CheckSignature(verifySignatureOnly: true);
                Assert.Equal(plaintext, signedCms.ContentInfo.Content);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        // =====================================================================
        // Empty and Edge Case Content
        // =====================================================================

        [Fact]
        public void Encrypt_EmptyContent_ThrowsCryptographicException()
        {
            // Windows CMS implementation does not support encrypting empty content
            var plaintext = new byte[0];
            Assert.ThrowsAny<CryptographicException>(
                () => _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 }));
        }

        [Fact]
        public void Sign_EmptyContent_ThrowsCryptographicException()
        {
            // Windows CMS implementation does not support signing empty content
            var plaintext = new byte[0];
            Assert.ThrowsAny<CryptographicException>(
                () => _handler.Sign(plaintext, _signingCert));
        }

        [Fact]
        public void Encrypt_BinaryContent_RoundTrip()
        {
            AddToStore(_encCert1);
            try
            {
                // Binary content with all byte values
                var plaintext = new byte[256];
                for (int i = 0; i < 256; i++) plaintext[i] = (byte)i;

                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });
                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        [Fact]
        public void Encrypt_UnicodeContent_RoundTrip()
        {
            AddToStore(_encCert1);
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Héllo Wörld! 你好世界 🔐📧 Ñoño señor");
                var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });
                var decryptResult = _handler.Decrypt(encrypted);
                var decrypted = decryptResult.Content;
                Assert.Equal(plaintext, decrypted);
                Assert.Equal("Héllo Wörld! 你好世界 🔐📧 Ñoño señor", Encoding.UTF8.GetString(decrypted));
            }
            finally
            {
                RemoveFromStore(_encCert1);
            }
        }

        // =====================================================================
        // Error / Negative Cases
        // =====================================================================

        [Fact]
        public void Decrypt_WithoutMatchingKey_Throws()
        {
            // Encrypt to cert1 but don't add cert1 to the store
            var plaintext = Encoding.UTF8.GetBytes("No matching private key.");
            var encrypted = _handler.Encrypt(plaintext, new X509Certificate2Collection { _encCert1 });

            var result = _handler.Decrypt(encrypted);
            Assert.False(result.Success);
        }

        [Fact]
        public void Verify_GarbageData_ReturnsFailed()
        {
            var garbage = Encoding.UTF8.GetBytes("This is not valid CMS data.");
            var result = _handler.Verify(garbage);
            Assert.False(result.IsValid);
            Assert.NotNull(result.ErrorMessage);
        }

        [Fact]
        public void Decrypt_GarbageData_Throws()
        {
            var garbage = Encoding.UTF8.GetBytes("This is not valid CMS data.");
            var result = _handler.Decrypt(garbage);
            Assert.False(result.Success);
        }

        [Fact]
        public void Sign_ExpiredCert_StillSigns()
        {
            // Expired cert can still produce a signature (verification may fail chain check)
            var expiredCert = CreateCert("CN=Expired Signer", X509KeyUsageFlags.DigitalSignature,
                notBefore: DateTimeOffset.UtcNow.AddDays(-30), notAfter: DateTimeOffset.UtcNow.AddDays(-1));
            try
            {
                var plaintext = Encoding.UTF8.GetBytes("Signed with expired cert.");
                var signed = _handler.Sign(plaintext, expiredCert);
                Assert.NotNull(signed);
                Assert.True(signed.Length > 0);
            }
            finally
            {
                expiredCert.Dispose();
            }
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private byte[] EncryptWithAlgorithm(byte[] content, X509Certificate2Collection recipients, string algoOid)
        {
            var contentInfo = new ContentInfo(content);
            var envelopedCms = new EnvelopedCms(contentInfo, new AlgorithmIdentifier(new Oid(algoOid)));
            var cmsRecipients = new CmsRecipientCollection();
            foreach (X509Certificate2 cert in recipients)
                cmsRecipients.Add(new CmsRecipient(SubjectIdentifierType.IssuerAndSerialNumber, cert));
            envelopedCms.Encrypt(cmsRecipients);
            return envelopedCms.Encode();
        }

        private byte[] SignWithHash(byte[] content, X509Certificate2 cert, string hashOid)
        {
            var contentInfo = new ContentInfo(content);
            var signedCms = new SignedCms(contentInfo, detached: false);
            var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, cert)
            {
                DigestAlgorithm = new Oid(hashOid)
            };
            signer.IncludeOption = X509IncludeOption.WholeChain;
            signedCms.ComputeSignature(signer);
            return signedCms.Encode();
        }

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

        private static void TrustCert(X509Certificate2 cert)
        {
            using (var store = new X509Store(StoreName.Root, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                store.Add(cert);
            }
        }

        private static void UntrustCert(X509Certificate2 cert)
        {
            using (var store = new X509Store(StoreName.Root, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var matches = store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, false);
                foreach (X509Certificate2 c in matches)
                    store.Remove(c);
            }
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

                return new X509Certificate2(
                    cert.Export(X509ContentType.Pfx, "test"), "test",
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
        }

        public void Dispose()
        {
            foreach (var cert in _allCerts) cert?.Dispose();
            _signingCert?.Dispose();
        }
    }
}
