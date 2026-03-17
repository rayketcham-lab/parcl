using System.Linq;
using Parcl.Core.Config;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    public class ParclSettingsTests
    {
        [Fact]
        public void CreateDefault_HasExpectedDefaults()
        {
            var settings = new ParclSettings();

            Assert.NotNull(settings.UserProfile);
            Assert.Empty(settings.LdapDirectories);
            Assert.Equal("AES-256-CBC", settings.Crypto.EncryptionAlgorithm);
            Assert.Equal("SHA-256", settings.Crypto.HashAlgorithm);
            Assert.False(settings.Crypto.AlwaysSign);
            Assert.False(settings.Crypto.AlwaysEncrypt);
            Assert.Equal(CertValidationMode.Relaxed, settings.Crypto.ValidationMode);
            Assert.True(settings.Cache.EnableCertCache);
            Assert.Equal(24, settings.Cache.CacheExpirationHours);
            Assert.Equal(500, settings.Cache.MaxCacheEntries);
            Assert.Equal(LookupTrigger.OnCompose, settings.Behavior.AutoLookup);
            Assert.True(settings.Behavior.PromptOnMissingCert);
        }

        [Fact]
        public void LdapDirectoryEntry_HasExpectedDefaults()
        {
            var entry = new LdapDirectoryEntry();

            Assert.Equal(636, entry.Port);
            Assert.True(entry.UseSsl);
            Assert.Equal("(mail={0})", entry.SearchFilter);
            Assert.Equal("userCertificate;binary", entry.CertAttribute);
            Assert.Equal(AuthType.Negotiate, entry.AuthType);
            Assert.True(entry.Enabled);
        }

        [Fact]
        public void LdapDirectoryEntry_ConnectionString_FormatsCorrectly()
        {
            var entry = new LdapDirectoryEntry
            {
                Server = "ldap.example.com",
                Port = 636,
                UseSsl = true
            };

            Assert.Equal("ldaps://ldap.example.com:636", entry.ConnectionString);
        }

        [Fact]
        public void UseNativeSmime_DefaultsTrue()
        {
            var settings = new ParclSettings();
            Assert.True(settings.Crypto.UseNativeSmime,
                "UseNativeSmime should default to true (native S/MIME for interoperability). " +
                "Changing this default breaks encryption for recipients without Parcl.");
        }

        [Fact]
        public void OpaqueSign_DefaultsFalse()
        {
            var settings = new ParclSettings();
            Assert.False(settings.Crypto.OpaqueSign);
        }

        [Fact]
        public void IncludeCertChain_DefaultsTrue()
        {
            var settings = new ParclSettings();
            Assert.True(settings.Crypto.IncludeCertChain);
        }

        [Fact]
        public void AlwaysEncrypt_DefaultsFalse()
        {
            var settings = new ParclSettings();
            Assert.False(settings.Crypto.AlwaysEncrypt);
        }

        [Fact]
        public void AlwaysSign_DefaultsFalse()
        {
            var settings = new ParclSettings();
            Assert.False(settings.Crypto.AlwaysSign);
        }

        [Fact]
        public void LdapDirectoryEntry_ConnectionString_NoSsl()
        {
            var entry = new LdapDirectoryEntry
            {
                Server = "ldap.example.com",
                Port = 389,
                UseSsl = false
            };

            Assert.Equal("ldap://ldap.example.com:389", entry.ConnectionString);
        }
    }
}
