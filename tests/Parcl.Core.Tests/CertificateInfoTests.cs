using System;
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
    }
}
