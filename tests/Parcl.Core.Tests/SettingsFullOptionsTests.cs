using System;
using System.IO;
using Parcl.Core.Config;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// E2E tests for all settings options — crypto prefs, behavior, cache, LDAP configs.
    /// </summary>
    [Collection("Settings")]
    public class SettingsFullOptionsTests : IDisposable
    {
        private readonly string _settingsDir;
        private readonly string _settingsFile;
        private readonly string _backupFile;

        public SettingsFullOptionsTests()
        {
            _settingsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl");
            _settingsFile = Path.Combine(_settingsDir, "settings.json");
            _backupFile = Path.Combine(_settingsDir, "settings.json.testbackup2");

            if (File.Exists(_settingsFile))
                File.Copy(_settingsFile, _backupFile, overwrite: true);
        }

        // =====================================================================
        // Encryption Algorithm Options
        // =====================================================================

        [Theory]
        [InlineData("AES-128-CBC")]
        [InlineData("AES-192-CBC")]
        [InlineData("AES-256-CBC")]
        public void CryptoSettings_EncryptionAlgorithm_PersistsAllValues(string algo)
        {
            var settings = new ParclSettings();
            settings.Crypto.EncryptionAlgorithm = algo;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(algo, loaded.Crypto.EncryptionAlgorithm);
        }

        // =====================================================================
        // Hash Algorithm Options
        // =====================================================================

        [Theory]
        [InlineData("SHA-256")]
        [InlineData("SHA-384")]
        [InlineData("SHA-512")]
        public void CryptoSettings_HashAlgorithm_PersistsAllValues(string hash)
        {
            var settings = new ParclSettings();
            settings.Crypto.HashAlgorithm = hash;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(hash, loaded.Crypto.HashAlgorithm);
        }

        // =====================================================================
        // Always Sign / Always Encrypt Toggles
        // =====================================================================

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void CryptoSettings_AlwaysSignEncrypt_AllCombinations(bool alwaysSign, bool alwaysEncrypt)
        {
            var settings = new ParclSettings();
            settings.Crypto.AlwaysSign = alwaysSign;
            settings.Crypto.AlwaysEncrypt = alwaysEncrypt;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(alwaysSign, loaded.Crypto.AlwaysSign);
            Assert.Equal(alwaysEncrypt, loaded.Crypto.AlwaysEncrypt);
        }

        // =====================================================================
        // Behavior: Lookup Trigger
        // =====================================================================

        [Theory]
        [InlineData(LookupTrigger.Manual)]
        [InlineData(LookupTrigger.OnCompose)]
        [InlineData(LookupTrigger.OnSend)]
        public void BehaviorSettings_AutoLookup_AllModes(LookupTrigger trigger)
        {
            var settings = new ParclSettings();
            settings.Behavior.AutoLookup = trigger;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(trigger, loaded.Behavior.AutoLookup);
        }

        [Theory]
        [InlineData(true, true)]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        public void BehaviorSettings_PromptAndStatus_AllCombinations(bool prompt, bool status)
        {
            var settings = new ParclSettings();
            settings.Behavior.PromptOnMissingCert = prompt;
            settings.Behavior.ShowStatusBar = status;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(prompt, loaded.Behavior.PromptOnMissingCert);
            Assert.Equal(status, loaded.Behavior.ShowStatusBar);
        }

        // =====================================================================
        // Cache Options
        // =====================================================================

        [Fact]
        public void CacheSettings_EnableDisable()
        {
            var settings = new ParclSettings();
            settings.Cache.EnableCertCache = false;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.False(loaded.Cache.EnableCertCache);

            loaded.Cache.EnableCertCache = true;
            loaded.Save();

            var reloaded = ParclSettings.Load();
            Assert.True(reloaded.Cache.EnableCertCache);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(24)]
        [InlineData(168)]
        [InlineData(720)]
        public void CacheSettings_ExpirationHours_AllValues(int hours)
        {
            var settings = new ParclSettings();
            settings.Cache.CacheExpirationHours = hours;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(hours, loaded.Cache.CacheExpirationHours);
        }

        [Theory]
        [InlineData(10)]
        [InlineData(500)]
        [InlineData(5000)]
        [InlineData(10000)]
        public void CacheSettings_MaxEntries_AllValues(int max)
        {
            var settings = new ParclSettings();
            settings.Cache.MaxCacheEntries = max;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(max, loaded.Cache.MaxCacheEntries);
        }

        // =====================================================================
        // LDAP Directory Configuration
        // =====================================================================

        [Fact]
        public void LdapSettings_MultipleDirectories_PersistAll()
        {
            var settings = new ParclSettings();
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Corp AD",
                Server = "ldap.corp.com",
                Port = 389,
                UseSsl = false,
                BaseDn = "dc=corp,dc=com",
                AuthType = AuthType.Negotiate,
                Enabled = true
            });
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Partner LDAP",
                Server = "ldap.partner.org",
                Port = 636,
                UseSsl = true,
                BaseDn = "ou=users,dc=partner,dc=org",
                AuthType = AuthType.Simple,
                BindDn = "cn=reader,dc=partner,dc=org",
                Enabled = true
            });
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Public Directory",
                Server = "ldap.public.example",
                Port = 389,
                UseSsl = false,
                BaseDn = "dc=public,dc=example",
                AuthType = AuthType.Anonymous,
                Enabled = false
            });
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal(3, loaded.LdapDirectories.Count);

            // Corp AD
            Assert.Equal("Corp AD", loaded.LdapDirectories[0].Name);
            Assert.Equal(389, loaded.LdapDirectories[0].Port);
            Assert.False(loaded.LdapDirectories[0].UseSsl);
            Assert.Equal(AuthType.Negotiate, loaded.LdapDirectories[0].AuthType);
            Assert.Equal("ldap://ldap.corp.com:389", loaded.LdapDirectories[0].ConnectionString);

            // Partner LDAP
            Assert.Equal("Partner LDAP", loaded.LdapDirectories[1].Name);
            Assert.Equal(636, loaded.LdapDirectories[1].Port);
            Assert.True(loaded.LdapDirectories[1].UseSsl);
            Assert.Equal(AuthType.Simple, loaded.LdapDirectories[1].AuthType);
            Assert.Equal("cn=reader,dc=partner,dc=org", loaded.LdapDirectories[1].BindDn);
            Assert.Equal("ldaps://ldap.partner.org:636", loaded.LdapDirectories[1].ConnectionString);

            // Public Directory (disabled)
            Assert.Equal("Public Directory", loaded.LdapDirectories[2].Name);
            Assert.Equal(AuthType.Anonymous, loaded.LdapDirectories[2].AuthType);
            Assert.False(loaded.LdapDirectories[2].Enabled);
        }

        [Theory]
        [InlineData(AuthType.Anonymous)]
        [InlineData(AuthType.Simple)]
        [InlineData(AuthType.Negotiate)]
        public void LdapSettings_AuthType_AllValues(AuthType authType)
        {
            var settings = new ParclSettings();
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Test",
                Server = "ldap.test.com",
                AuthType = authType
            });
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Single(loaded.LdapDirectories);
            Assert.Equal(authType, loaded.LdapDirectories[0].AuthType);
        }

        [Fact]
        public void LdapSettings_CustomSearchFilter_Persists()
        {
            var settings = new ParclSettings();
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Custom Filter",
                Server = "ldap.custom.com",
                SearchFilter = "(&(objectClass=person)(mail={0}))",
                CertAttribute = "userSMIMECertificate;binary"
            });
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal("(&(objectClass=person)(mail={0}))", loaded.LdapDirectories[0].SearchFilter);
            Assert.Equal("userSMIMECertificate;binary", loaded.LdapDirectories[0].CertAttribute);
        }

        [Fact]
        public void LdapSettings_SslToggle_AffectsConnectionString()
        {
            var entry = new LdapDirectoryEntry
            {
                Server = "ldap.example.com",
                Port = 636,
                UseSsl = true
            };
            Assert.Equal("ldaps://ldap.example.com:636", entry.ConnectionString);

            entry.UseSsl = false;
            Assert.Equal("ldap://ldap.example.com:636", entry.ConnectionString);
        }

        // =====================================================================
        // User Profile
        // =====================================================================

        [Fact]
        public void UserProfile_AllFields_Persist()
        {
            var settings = new ParclSettings();
            settings.UserProfile = new UserProfile
            {
                EmailAddress = "ray@example.com",
                DisplayName = "Ray Ketcham",
                SigningCertThumbprint = "AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555",
                EncryptionCertThumbprint = "FFFF6666777788889999AAAA0000BBBBCCCCDDDD"
            };
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal("ray@example.com", loaded.UserProfile.EmailAddress);
            Assert.Equal("Ray Ketcham", loaded.UserProfile.DisplayName);
            Assert.Equal("AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555", loaded.UserProfile.SigningCertThumbprint);
            Assert.Equal("FFFF6666777788889999AAAA0000BBBBCCCCDDDD", loaded.UserProfile.EncryptionCertThumbprint);
        }

        [Fact]
        public void UserProfile_NullThumbprints_Allowed()
        {
            var settings = new ParclSettings();
            settings.UserProfile.SigningCertThumbprint = null;
            settings.UserProfile.EncryptionCertThumbprint = null;
            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Null(loaded.UserProfile.SigningCertThumbprint);
            Assert.Null(loaded.UserProfile.EncryptionCertThumbprint);
        }

        // =====================================================================
        // Full Settings Combo
        // =====================================================================

        [Fact]
        public void FullSettings_AllOptionsSet_PersistsCorrectly()
        {
            var settings = new ParclSettings
            {
                UserProfile = new UserProfile
                {
                    EmailAddress = "full@test.com",
                    DisplayName = "Full Test",
                    SigningCertThumbprint = "SIGNING1234",
                    EncryptionCertThumbprint = "ENCRYPT5678"
                },
                Crypto = new CryptoPreferences
                {
                    EncryptionAlgorithm = "AES-128-CBC",
                    HashAlgorithm = "SHA-512",
                    AlwaysSign = true,
                    AlwaysEncrypt = true
                },
                Cache = new CacheSettings
                {
                    EnableCertCache = false,
                    CacheExpirationHours = 168,
                    MaxCacheEntries = 1000
                },
                Behavior = new BehaviorSettings
                {
                    AutoLookup = LookupTrigger.OnSend,
                    PromptOnMissingCert = false,
                    ShowStatusBar = false
                }
            };
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Full Test LDAP",
                Server = "ldap.fulltest.com",
                Port = 636,
                UseSsl = true,
                BaseDn = "dc=fulltest,dc=com",
                SearchFilter = "(&(mail={0})(objectClass=user))",
                CertAttribute = "userSMIMECertificate;binary",
                AuthType = AuthType.Simple,
                BindDn = "cn=app,dc=fulltest,dc=com",
                Enabled = true
            });
            settings.Save();

            var loaded = ParclSettings.Load();

            // User profile
            Assert.Equal("full@test.com", loaded.UserProfile.EmailAddress);

            // Crypto
            Assert.Equal("AES-128-CBC", loaded.Crypto.EncryptionAlgorithm);
            Assert.Equal("SHA-512", loaded.Crypto.HashAlgorithm);
            Assert.True(loaded.Crypto.AlwaysSign);
            Assert.True(loaded.Crypto.AlwaysEncrypt);

            // Cache
            Assert.False(loaded.Cache.EnableCertCache);
            Assert.Equal(168, loaded.Cache.CacheExpirationHours);
            Assert.Equal(1000, loaded.Cache.MaxCacheEntries);

            // Behavior
            Assert.Equal(LookupTrigger.OnSend, loaded.Behavior.AutoLookup);
            Assert.False(loaded.Behavior.PromptOnMissingCert);
            Assert.False(loaded.Behavior.ShowStatusBar);

            // LDAP
            Assert.Single(loaded.LdapDirectories);
            Assert.Equal("ldaps://ldap.fulltest.com:636", loaded.LdapDirectories[0].ConnectionString);
            Assert.Equal(AuthType.Simple, loaded.LdapDirectories[0].AuthType);
            Assert.Equal("cn=app,dc=fulltest,dc=com", loaded.LdapDirectories[0].BindDn);
        }

        public void Dispose()
        {
            if (File.Exists(_backupFile))
            {
                File.Copy(_backupFile, _settingsFile, overwrite: true);
                File.Delete(_backupFile);
            }
        }
    }
}
