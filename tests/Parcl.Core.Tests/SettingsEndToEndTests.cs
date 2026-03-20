using System;
using System.IO;
using Parcl.Core.Config;
using Parcl.Core.Models;
using Xunit;

namespace Parcl.Core.Tests
{
    [Collection("Settings")]
    public class SettingsEndToEndTests : IDisposable
    {
        private readonly string _settingsDir;
        private readonly string _settingsFile;
        private readonly string _backupFile;

        public SettingsEndToEndTests()
        {
            _settingsDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl");
            _settingsFile = Path.Combine(_settingsDir, "settings.json");
            _backupFile = Path.Combine(_settingsDir, "settings.json.testbackup");

            if (File.Exists(_settingsFile))
                File.Copy(_settingsFile, _backupFile, overwrite: true);
        }

        [Fact]
        public void SaveAndLoad_RoundTrip()
        {
            var settings = new ParclSettings
            {
                UserProfile = new UserProfile
                {
                    EmailAddress = "ray@example.com",
                    DisplayName = "Ray Ketcham",
                    SigningCertThumbprint = "ABCD1234"
                },
                Crypto = new CryptoPreferences
                {
                    AlwaysSign = true,
                    AlwaysEncrypt = true,
                    EncryptionAlgorithm = "AES-256-CBC",
                    HashAlgorithm = "SHA-256"
                }
            };
            settings.LdapDirectories.Add(new LdapDirectoryEntry
            {
                Name = "Corp LDAP",
                Server = "ldap.corp.example.com",
                Port = 636,
                UseSsl = true,
                BaseDn = "dc=corp,dc=example,dc=com"
            });

            settings.Save();

            var loaded = ParclSettings.Load();
            Assert.Equal("ray@example.com", loaded.UserProfile.EmailAddress);
            Assert.Equal("Ray Ketcham", loaded.UserProfile.DisplayName);
            Assert.Equal("ABCD1234", loaded.UserProfile.SigningCertThumbprint);
            Assert.True(loaded.Crypto.AlwaysSign);
            Assert.True(loaded.Crypto.AlwaysEncrypt);
            Assert.Single(loaded.LdapDirectories);
            Assert.Equal("Corp LDAP", loaded.LdapDirectories[0].Name);
            Assert.Equal(636, loaded.LdapDirectories[0].Port);
            Assert.True(loaded.LdapDirectories[0].UseSsl);
            Assert.Equal("ldaps://ldap.corp.example.com:636", loaded.LdapDirectories[0].ConnectionString);
        }

        [Fact]
        public void Load_WhenFileMissing_ReturnsDefaults()
        {
            // Delete settings to test default creation
            if (File.Exists(_settingsFile))
                File.Delete(_settingsFile);

            var settings = ParclSettings.Load();
            Assert.NotNull(settings);
            Assert.Equal("AES-256-CBC", settings.Crypto.EncryptionAlgorithm);
            Assert.Equal(LookupTrigger.OnCompose, settings.Behavior.AutoLookup);
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
