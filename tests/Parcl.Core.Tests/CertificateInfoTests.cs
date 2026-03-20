using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    public class CertificateInfoTests
    {
        [Fact]
        public void IsExpired_WhenNotAfterInPast_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                NotBefore = DateTime.UtcNow.AddYears(-2),
                NotAfter = DateTime.UtcNow.AddDays(-1)
            };

            Assert.True(info.IsExpired);
            Assert.False(info.IsValid);
        }

        [Fact]
        public void IsNotYetValid_WhenNotBeforeInFuture_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                NotBefore = DateTime.UtcNow.AddDays(1),
                NotAfter = DateTime.UtcNow.AddYears(1)
            };

            Assert.True(info.IsNotYetValid);
            Assert.False(info.IsValid);
        }

        [Fact]
        public void IsValid_WhenWithinValidityPeriod_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                NotBefore = DateTime.UtcNow.AddDays(-1),
                NotAfter = DateTime.UtcNow.AddYears(1)
            };

            Assert.True(info.IsValid);
            Assert.False(info.IsExpired);
            Assert.False(info.IsNotYetValid);
        }

        [Fact]
        public void IsSigningCert_WithDigitalSignatureUsage_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                KeyUsage = X509KeyUsageFlags.DigitalSignature
            };

            Assert.True(info.IsSigningCert);
            Assert.False(info.IsEncryptionCert);
        }

        [Fact]
        public void IsEncryptionCert_WithKeyEnciphermentUsage_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                KeyUsage = X509KeyUsageFlags.KeyEncipherment
            };

            Assert.True(info.IsEncryptionCert);
            Assert.False(info.IsSigningCert);
        }

        [Fact]
        public void IsSigningCert_WithNonRepudiationUsage_ReturnsTrue()
        {
            var info = new CertificateInfo
            {
                KeyUsage = X509KeyUsageFlags.NonRepudiation
            };

            Assert.True(info.IsSigningCert);
        }

        [Fact]
        public void ToString_IncludesSubjectAndThumbprint()
        {
            var info = new CertificateInfo
            {
                Subject = "CN=Test User",
                Thumbprint = "ABCDEF1234567890ABCDEF1234567890ABCDEF12",
                NotAfter = new DateTime(2027, 12, 31)
            };

            var str = info.ToString();
            Assert.Contains("CN=Test User", str);
            Assert.Contains("ABCDEF12", str);
            Assert.Contains("2027-12-31", str);
        }
        // =====================================================================
        // FromX509 — SAN Email Extraction
        // =====================================================================

        [Fact]
        public void FromX509_SimpleRfc822San_ExtractsEmail()
        {
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest("CN=Simple SAN Test", rsa,
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var sanBuilder = new SubjectAlternativeNameBuilder();
            sanBuilder.AddEmailAddress("user@example.com");
            req.CertificateExtensions.Add(sanBuilder.Build());
            req.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.DigitalSignature, true));
            using var cert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));

            var info = CertificateInfo.FromX509(cert);

            Assert.Equal("user@example.com", info.Email);
        }

        [Fact]
        public void FromX509_MultiSan_ExtractsRfc822NotPrincipalName()
        {
            // Simulates enterprise certs (e.g., RTX) with Principal Name before RFC822 Name.
            // The old code extracted the Principal Name email instead of RFC822.
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest("CN=Enterprise User", rsa,
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var sanBuilder = new SubjectAlternativeNameBuilder();
            // Add UPN first (like enterprise certs do)
            sanBuilder.AddUserPrincipalName("E21127560@adxuser.com");
            // Then the actual RFC822 email
            sanBuilder.AddEmailAddress("james.r.ketcham@rtx.com");
            req.CertificateExtensions.Add(sanBuilder.Build());
            req.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.KeyEncipherment, true));
            using var cert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));

            var info = CertificateInfo.FromX509(cert);

            Assert.Equal("james.r.ketcham@rtx.com", info.Email);
        }

        [Fact]
        public void FromX509_NoSan_ReturnsEmptyEmail()
        {
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest("CN=No SAN Cert", rsa,
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            req.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.DigitalSignature, true));
            using var cert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));

            var info = CertificateInfo.FromX509(cert);

            Assert.Equal(string.Empty, info.Email);
        }

        [Fact]
        public void FromX509_SanWithoutRfc822_ReturnsEmptyEmail()
        {
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest("CN=DNS Only SAN", rsa,
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var sanBuilder = new SubjectAlternativeNameBuilder();
            sanBuilder.AddDnsName("mail.example.com");
            req.CertificateExtensions.Add(sanBuilder.Build());
            req.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.KeyEncipherment, true));
            using var cert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));

            var info = CertificateInfo.FromX509(cert);

            Assert.Equal(string.Empty, info.Email);
        }

        [Fact]
        public void FromX509_MultipleRfc822Sans_ReturnsFirst()
        {
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest("CN=Multi RFC822", rsa,
                HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var sanBuilder = new SubjectAlternativeNameBuilder();
            sanBuilder.AddEmailAddress("primary@example.com");
            sanBuilder.AddEmailAddress("secondary@example.com");
            req.CertificateExtensions.Add(sanBuilder.Build());
            req.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.KeyEncipherment, true));
            using var cert = req.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddHours(1));

            var info = CertificateInfo.FromX509(cert);

            Assert.Equal("primary@example.com", info.Email);
        }
    }
}
