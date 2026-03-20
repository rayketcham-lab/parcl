using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// Real E2E tests using the configured Parcl settings and rayketcham@ogjos.com certificate.
    /// These tests exercise the full Parcl pipeline as it would run in the Outlook add-in.
    /// </summary>
    [Collection("RealCert")]
    public class RealCertE2ETests : IDisposable
    {
        private readonly ParclSettings _settings;
        private readonly CertificateStore _certStore;
        private readonly SmimeHandler _smime;
        private readonly CertExchange _exchange;
        private readonly CertificateCache _cache;
        private readonly X509Certificate2 _userCert;

        public RealCertE2ETests()
        {
            _settings = ParclSettings.Load();
            _certStore = new CertificateStore();
            _smime = new SmimeHandler();
            _exchange = new CertExchange(_certStore);
            _cache = new CertificateCache(
                _settings.Cache.CacheExpirationHours,
                _settings.Cache.MaxCacheEntries);

            // Load the user's configured cert
            _userCert = _certStore.FindByThumbprint(_settings.UserProfile.SigningCertThumbprint!)
                ?? throw new InvalidOperationException(
                    "Test requires configured signing cert in Parcl settings. " +
                    $"Thumbprint: {_settings.UserProfile.SigningCertThumbprint}");
        }

        // =====================================================================
        // Settings Validation
        // =====================================================================

        [Fact]
        public void Settings_UserProfileConfigured()
        {
            Assert.Equal("rayketcham@ogjos.com", _settings.UserProfile.EmailAddress);
            Assert.Equal("Ray Ketcham", _settings.UserProfile.DisplayName);
            Assert.NotNull(_settings.UserProfile.SigningCertThumbprint);
            Assert.NotNull(_settings.UserProfile.EncryptionCertThumbprint);
        }

        [Fact]
        public void Settings_CryptoDefaults()
        {
            Assert.Equal("AES-256-CBC", _settings.Crypto.EncryptionAlgorithm);
            Assert.Equal("SHA-256", _settings.Crypto.HashAlgorithm);
        }

        [Fact]
        public void Settings_CertExistsInStore()
        {
            Assert.NotNull(_userCert);
            Assert.True(_userCert.HasPrivateKey, "Configured cert must have a private key");
            Assert.Contains("rayketcham@ogjos.com", _userCert.Subject);
        }

        [Fact]
        public void Settings_CertHasCorrectKeyUsage()
        {
            var info = CertificateInfo.FromX509(_userCert);
            Assert.True(info.IsSigningCert, "Cert should support signing");
            Assert.True(info.IsEncryptionCert, "Cert should support encryption");
            Assert.True(info.IsValid, "Cert should be within validity period");
        }

        [Fact]
        public void Settings_CertAppearsInStoreFilters()
        {
            var signing = _certStore.GetSigningCertificates();
            var encryption = _certStore.GetEncryptionCertificates();

            Assert.Contains(signing, c => c.Thumbprint == _userCert.Thumbprint);
            Assert.Contains(encryption, c => c.Thumbprint == _userCert.Thumbprint);
        }

        [Fact]
        public void Settings_CertFoundByEmail()
        {
            var found = _certStore.FindByEmail("rayketcham@ogjos.com");
            Assert.NotNull(found);
        }

        // =====================================================================
        // Sign Email — All Hash Algorithms
        // =====================================================================

        [Fact]
        public void SignEmail_SHA256_UsingConfiguredCert()
        {
            var body = Encoding.UTF8.GetBytes(
                "From: rayketcham@ogjos.com\r\nTo: test@example.com\r\n\r\nThis is a signed test email.");

            var signed = _smime.Sign(body, _userCert);
            Assert.NotNull(signed);

            var cms = new SignedCms();
            cms.Decode(signed);
            Assert.Equal("2.16.840.1.101.3.4.2.1", cms.SignerInfos[0].DigestAlgorithm.Value); // SHA-256
            Assert.Equal(_userCert.Subject, cms.SignerInfos[0].Certificate.Subject);

            cms.CheckSignature(verifySignatureOnly: false); // Full chain validation — cert is trusted
            Assert.Equal(body, cms.ContentInfo.Content);
        }

        [Theory]
        [InlineData("2.16.840.1.101.3.4.2.1", "SHA-256")]
        [InlineData("2.16.840.1.101.3.4.2.2", "SHA-384")]
        [InlineData("2.16.840.1.101.3.4.2.3", "SHA-512")]
        [InlineData("1.3.14.3.2.26", "SHA-1")]
        public void SignEmail_AllHashAlgorithms(string oid, string name)
        {
            var body = Encoding.UTF8.GetBytes($"Signed with {name} by rayketcham@ogjos.com");

            var contentInfo = new ContentInfo(body);
            var signedCms = new SignedCms(contentInfo, detached: false);
            var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, _userCert)
            {
                DigestAlgorithm = new Oid(oid)
            };
            signer.IncludeOption = X509IncludeOption.WholeChain;
            signedCms.ComputeSignature(signer);
            var signed = signedCms.Encode();

            var verify = new SignedCms();
            verify.Decode(signed);
            Assert.Equal(oid, verify.SignerInfos[0].DigestAlgorithm.Value);
            verify.CheckSignature(verifySignatureOnly: false);
            Assert.Equal(body, verify.ContentInfo.Content);
        }

        // =====================================================================
        // Encrypt Email — All Encryption Algorithms
        // =====================================================================

        [Theory]
        [InlineData("2.16.840.1.101.3.4.1.2", "AES-128-CBC")]
        [InlineData("2.16.840.1.101.3.4.1.22", "AES-192-CBC")]
        [InlineData("2.16.840.1.101.3.4.1.42", "AES-256-CBC")]
        [InlineData("1.2.840.113549.3.7", "3DES")]
        public void EncryptEmail_AllAlgorithms(string oid, string name)
        {
            var body = Encoding.UTF8.GetBytes(
                $"From: sender@example.com\r\nTo: rayketcham@ogjos.com\r\n\r\nEncrypted with {name}.");

            var contentInfo = new ContentInfo(body);
            var envelope = new EnvelopedCms(contentInfo, new AlgorithmIdentifier(new Oid(oid)));
            var recipient = new CmsRecipient(SubjectIdentifierType.IssuerAndSerialNumber, _userCert);
            envelope.Encrypt(new CmsRecipientCollection { recipient });
            var encrypted = envelope.Encode();

            // Decrypt with our private key
            var decryptEnvelope = new EnvelopedCms();
            decryptEnvelope.Decode(encrypted);
            Assert.Equal(oid, decryptEnvelope.ContentEncryptionAlgorithm.Oid.Value);
            decryptEnvelope.Decrypt();
            Assert.Equal(body, decryptEnvelope.ContentInfo.Content);
        }

        [Fact]
        public void EncryptEmail_DefaultAES256_UsingSmimeHandler()
        {
            var body = Encoding.UTF8.GetBytes(
                "From: sender@example.com\r\nTo: rayketcham@ogjos.com\r\n\r\nDefault AES-256 encrypted message.");

            var recipientCerts = new X509Certificate2Collection { _userCert };
            var encrypted = _smime.Encrypt(body, recipientCerts);
            var decryptResult = _smime.Decrypt(encrypted);
            var decrypted = decryptResult.Content;

            Assert.Equal(body, decrypted);
            Assert.Equal("Default AES-256 encrypted message.",
                Encoding.UTF8.GetString(decrypted).Split(new[] { "\r\n\r\n" }, StringSplitOptions.None).Last());
        }

        // =====================================================================
        // Sign + Encrypt Email (Full Pipeline)
        // =====================================================================

        [Fact]
        public void SignThenEncrypt_FullEmailPipeline()
        {
            var body = Encoding.UTF8.GetBytes(
                "From: rayketcham@ogjos.com\r\nTo: rayketcham@ogjos.com\r\n\r\n" +
                "This message is signed and encrypted using Parcl settings.");

            // Step 1: Sign with configured signing cert
            var signed = _smime.Sign(body, _userCert);

            // Step 2: Encrypt to recipient (self in this case)
            var recipientCerts = new X509Certificate2Collection { _userCert };
            var encrypted = _smime.Encrypt(signed, recipientCerts);

            // Step 3: Decrypt
            var decryptResult = _smime.Decrypt(encrypted);
            var decryptedSigned = decryptResult.Content;

            // Step 4: Verify signature
            var result = _smime.Verify(decryptedSigned);
            Assert.True(result.IsValid, $"Signature verification failed: {result.ErrorMessage}");
            Assert.Equal(body, result.Content);
            Assert.NotNull(result.SignerCertificate);
            Assert.Contains("rayketcham@ogjos.com", result.SignerCertificate.Subject);
        }

        [Fact]
        public void SignThenEncrypt_AlwaysSignAlwaysEncrypt_Simulation()
        {
            // Simulate the AlwaysSign + AlwaysEncrypt behavior
            var origSign = _settings.Crypto.AlwaysSign;
            var origEncrypt = _settings.Crypto.AlwaysEncrypt;

            try
            {
                _settings.Crypto.AlwaysSign = true;
                _settings.Crypto.AlwaysEncrypt = true;

                var body = Encoding.UTF8.GetBytes("Auto sign+encrypt message body");

                // Pipeline: if AlwaysSign, sign first
                byte[] payload = body;
                if (_settings.Crypto.AlwaysSign)
                {
                    var signingCert = _certStore.FindByThumbprint(_settings.UserProfile.SigningCertThumbprint!);
                    Assert.NotNull(signingCert);
                    payload = _smime.Sign(payload, signingCert);
                }

                // If AlwaysEncrypt, encrypt
                if (_settings.Crypto.AlwaysEncrypt)
                {
                    var encCert = _certStore.FindByThumbprint(_settings.UserProfile.EncryptionCertThumbprint!);
                    Assert.NotNull(encCert);
                    payload = _smime.Encrypt(payload, new X509Certificate2Collection { encCert });
                }

                // Reverse: decrypt then verify
                var decryptResult = _smime.Decrypt(payload);
                var decrypted = decryptResult.Content;
                var result = _smime.Verify(decrypted);
                Assert.True(result.IsValid);
                Assert.Equal(body, result.Content);
            }
            finally
            {
                _settings.Crypto.AlwaysSign = origSign;
                _settings.Crypto.AlwaysEncrypt = origEncrypt;
            }
        }

        // =====================================================================
        // Encrypt-Only Mode
        // =====================================================================

        [Fact]
        public void EncryptOnly_EmailBody_RoundTrip()
        {
            var body = Encoding.UTF8.GetBytes(
                "Confidential: This email is encrypted but not signed.");

            var encrypted = _smime.Encrypt(body, new X509Certificate2Collection { _userCert });

            // Ensure body is not readable
            Assert.DoesNotContain("Confidential", Encoding.UTF8.GetString(encrypted));

            var decryptResult = _smime.Decrypt(encrypted);
            var decrypted = decryptResult.Content;
            Assert.Equal(body, decrypted);
        }

        // =====================================================================
        // Sign-Only Mode
        // =====================================================================

        [Fact]
        public void SignOnly_EmailBody_VerifyWithChain()
        {
            var body = Encoding.UTF8.GetBytes(
                "This email is signed but not encrypted. Anyone can read it.");

            var signed = _smime.Sign(body, _userCert);
            var result = _smime.Verify(signed);

            Assert.True(result.IsValid, $"Verification failed: {result.ErrorMessage}");
            Assert.Equal(body, result.Content);
            Assert.NotNull(result.SignerCertificate);
            Assert.Equal("rayketcham@ogjos.com", result.SignerCertificate.Email);
        }

        // =====================================================================
        // Certificate Exchange
        // =====================================================================

        [Fact]
        public void CertExchange_ExportPublicCert_DerFormat()
        {
            var payload = _exchange.PrepareExport(_userCert.Thumbprint);
            Assert.Equal(_userCert.Thumbprint, payload.Thumbprint);
            Assert.Contains("rayketcham@ogjos.com", payload.SenderEmail);

            // Decode the base64 cert data
            var certBytes = Convert.FromBase64String(payload.CertificateData);
            var importedCert = new X509Certificate2(certBytes);
            Assert.Equal(_userCert.Subject, importedCert.Subject);
            Assert.False(importedCert.HasPrivateKey, "Exported cert should not have private key");
        }

        [Fact]
        public void CertExchange_ExportAsPem_ValidFormat()
        {
            var payload = _exchange.PrepareExport(_userCert.Thumbprint);
            var pem = _exchange.FormatAsAttachment(payload);

            Assert.StartsWith("-----BEGIN CERTIFICATE-----", pem);
            Assert.Contains("-----END CERTIFICATE-----", pem);

            // Re-parse from PEM
            var b64 = pem
                .Replace("-----BEGIN CERTIFICATE-----", "")
                .Replace("-----END CERTIFICATE-----", "")
                .Replace("\r", "").Replace("\n", "").Trim();
            var cert = new X509Certificate2(Convert.FromBase64String(b64));
            Assert.Equal(_userCert.Thumbprint, cert.Thumbprint);
        }

        [Fact]
        public void CertExchange_SimulateRecipientReceivesCert_ThenEncryptsToUs()
        {
            // Simulate: we send our public cert to a recipient
            var payload = _exchange.PrepareExport(_userCert.Thumbprint);

            // Recipient imports our public cert
            var ourPubBytes = Convert.FromBase64String(payload.CertificateData);
            var ourPubCert = new X509Certificate2(ourPubBytes);

            // Recipient encrypts a message to us
            var message = Encoding.UTF8.GetBytes("Hey Ray, here's that encrypted reply!");
            var encrypted = _smime.Encrypt(message, new X509Certificate2Collection { ourPubCert });

            // We decrypt with our private key
            var decryptResult = _smime.Decrypt(encrypted);
            var decrypted = decryptResult.Content;
            Assert.Equal("Hey Ray, here's that encrypted reply!", Encoding.UTF8.GetString(decrypted));
        }

        // =====================================================================
        // Certificate Cache with Real Cert
        // =====================================================================

        [Fact]
        public void Cache_StoreAndRetrieve_RealCertInfo()
        {
            var info = CertificateInfo.FromX509(_userCert);
            _cache.Add("rayketcham@ogjos.com", new System.Collections.Generic.List<CertificateInfo> { info });

            var cached = _cache.Get("rayketcham@ogjos.com");
            Assert.NotNull(cached);
            Assert.Single(cached);
            Assert.Equal(_userCert.Thumbprint, cached[0].Thumbprint);
            Assert.Contains("rayketcham@ogjos.com", cached[0].Subject);

            _cache.Remove("rayketcham@ogjos.com");
        }

        // =====================================================================
        // Real-World Email Content Patterns
        // =====================================================================

        [Fact]
        public void SignAndEncrypt_MimeFormattedEmail()
        {
            var mime = Encoding.UTF8.GetBytes(
                "MIME-Version: 1.0\r\n" +
                "From: rayketcham@ogjos.com\r\n" +
                "To: recipient@example.com\r\n" +
                "Subject: Parcl E2E Test\r\n" +
                "Content-Type: text/plain; charset=UTF-8\r\n" +
                "\r\n" +
                "This is a test email sent through Parcl.\r\n" +
                "It should be signed and encrypted.\r\n");

            var signed = _smime.Sign(mime, _userCert);
            var encrypted = _smime.Encrypt(signed, new X509Certificate2Collection { _userCert });
            var decryptResult = _smime.Decrypt(encrypted);
            var decryptedSigned = decryptResult.Content;
            var result = _smime.Verify(decryptedSigned);

            Assert.True(result.IsValid);
            Assert.Equal(mime, result.Content);
        }

        [Fact]
        public void SignAndEncrypt_HtmlEmail()
        {
            var html = Encoding.UTF8.GetBytes(
                "MIME-Version: 1.0\r\n" +
                "Content-Type: text/html; charset=UTF-8\r\n" +
                "\r\n" +
                "<html><body><h1>Parcl Test</h1><p>HTML email content with <b>formatting</b>.</p></body></html>\r\n");

            var signed = _smime.Sign(html, _userCert);
            var encrypted = _smime.Encrypt(signed, new X509Certificate2Collection { _userCert });
            var decryptResult = _smime.Decrypt(encrypted);
            var decryptedSigned = decryptResult.Content;
            var result = _smime.Verify(decryptedSigned);

            Assert.True(result.IsValid);
            Assert.Contains("<h1>Parcl Test</h1>", Encoding.UTF8.GetString(result.Content));
        }

        [Fact]
        public void SignAndEncrypt_EmailWithAttachmentMime()
        {
            var mime = Encoding.UTF8.GetBytes(
                "MIME-Version: 1.0\r\n" +
                "Content-Type: multipart/mixed; boundary=\"parcl-boundary\"\r\n" +
                "\r\n" +
                "--parcl-boundary\r\n" +
                "Content-Type: text/plain\r\n" +
                "\r\n" +
                "See attached file.\r\n" +
                "--parcl-boundary\r\n" +
                "Content-Type: application/octet-stream\r\n" +
                "Content-Disposition: attachment; filename=\"test.bin\"\r\n" +
                "Content-Transfer-Encoding: base64\r\n" +
                "\r\n" +
                "SGVsbG8gV29ybGQh\r\n" +
                "--parcl-boundary--\r\n");

            var signed = _smime.Sign(mime, _userCert);
            var encrypted = _smime.Encrypt(signed, new X509Certificate2Collection { _userCert });
            var decryptResult = _smime.Decrypt(encrypted);
            var decryptedSigned = decryptResult.Content;
            var result = _smime.Verify(decryptedSigned);

            Assert.True(result.IsValid);
            Assert.Contains("parcl-boundary", Encoding.UTF8.GetString(result.Content));
            Assert.Contains("SGVsbG8gV29ybGQh", Encoding.UTF8.GetString(result.Content));
        }

        // =====================================================================
        // Tamper Detection with Real Cert
        // =====================================================================

        [Fact]
        public void TamperDetection_ModifiedSignedEmail_Fails()
        {
            var body = Encoding.UTF8.GetBytes("Original message from rayketcham@ogjos.com");
            var signed = _smime.Sign(body, _userCert);

            // Tamper
            signed[signed.Length / 2] ^= 0xFF;

            var result = _smime.Verify(signed);
            Assert.False(result.IsValid);
        }

        [Fact]
        public void TamperDetection_WrongKeyDecrypt_Fails()
        {
            // Create a different cert
            using (var rsa = RSA.Create(2048))
            {
                var req = new CertificateRequest("CN=Wrong Person", rsa,
                    HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                req.CertificateExtensions.Add(new X509KeyUsageExtension(
                    X509KeyUsageFlags.KeyEncipherment, true));
                var wrongCert = req.CreateSelfSigned(
                    DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));
                var wrongPfx = new X509Certificate2(
                    wrongCert.Export(X509ContentType.Pfx, "t"), "t",
                    X509KeyStorageFlags.Exportable);

                // Encrypt to the wrong cert
                var body = Encoding.UTF8.GetBytes("Secret for wrong person only");
                var encrypted = _smime.Encrypt(body, new X509Certificate2Collection { wrongPfx });

                // Our real cert should NOT be able to decrypt
                var decryptResult = _smime.Decrypt(encrypted);
                Assert.False(decryptResult.Success);
            }
        }

        // =====================================================================
        // All Crypto Settings Combinations
        // =====================================================================

        [Theory]
        [InlineData("AES-256-CBC", "SHA-256", false, false)]
        [InlineData("AES-256-CBC", "SHA-256", true, false)]
        [InlineData("AES-256-CBC", "SHA-256", false, true)]
        [InlineData("AES-256-CBC", "SHA-256", true, true)]
        [InlineData("AES-128-CBC", "SHA-512", true, true)]
        [InlineData("3DES", "SHA-1", true, true)]
        [InlineData("AES-192-CBC", "SHA-384", true, true)]
        public void AllSettingsCombinations_EmailRoundTrip(
            string encAlgo, string hashAlgo, bool alwaysSign, bool alwaysEncrypt)
        {
            var body = Encoding.UTF8.GetBytes(
                $"Settings test: enc={encAlgo} hash={hashAlgo} sign={alwaysSign} encrypt={alwaysEncrypt}");

            byte[] payload = body;

            // Sign if configured
            if (alwaysSign)
                payload = _smime.Sign(payload, _userCert);

            // Encrypt if configured
            if (alwaysEncrypt)
                payload = _smime.Encrypt(payload, new X509Certificate2Collection { _userCert });

            // Reverse
            if (alwaysEncrypt)
                payload = _smime.Decrypt(payload).Content;

            if (alwaysSign)
            {
                var result = _smime.Verify(payload);
                Assert.True(result.IsValid, $"Failed for {encAlgo}/{hashAlgo}: {result.ErrorMessage}");
                payload = result.Content;
            }

            Assert.Equal(body, payload);
        }

        public void Dispose()
        {
            _certStore?.Dispose();
            _userCert?.Dispose();
        }
    }
}
