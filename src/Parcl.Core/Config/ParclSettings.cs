using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using Parcl.Core.Models;

namespace Parcl.Core.Config
{
    public class ParclSettings
    {
        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl");

        private static readonly string SettingsFile =
            Path.Combine(SettingsDir, "settings.json");

        public UserProfile UserProfile { get; set; } = new UserProfile();
        public List<LdapDirectoryEntry> LdapDirectories { get; set; } = new List<LdapDirectoryEntry>();
        public CryptoPreferences Crypto { get; set; } = new CryptoPreferences();
        public CacheSettings Cache { get; set; } = new CacheSettings();
        public BehaviorSettings Behavior { get; set; } = new BehaviorSettings();

        public static ParclSettings Load()
        {
            if (!File.Exists(SettingsFile))
                return CreateDefault();

            var json = File.ReadAllText(SettingsFile);

            if (!SettingsIntegrity.VerifyHmac(json, out var failureReason))
            {
                Trace.TraceWarning($"[Parcl.Settings] {failureReason}");
            }

            return JsonConvert.DeserializeObject<ParclSettings>(json) ?? CreateDefault();
        }

        public void Save()
        {
            Directory.CreateDirectory(SettingsDir);
            var json = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(SettingsFile, json);
            SettingsIntegrity.WriteHmac(json);
        }

        public void Reload()
        {
            var fresh = Load();
            UserProfile = fresh.UserProfile;
            LdapDirectories = fresh.LdapDirectories;
            Crypto = fresh.Crypto;
            Cache = fresh.Cache;
            Behavior = fresh.Behavior;
        }

        private static ParclSettings CreateDefault()
        {
            var settings = new ParclSettings();
            settings.Save();
            return settings;
        }
    }

    public class CryptoPreferences
    {
        public string EncryptionAlgorithm { get; set; } = "AES-256-CBC";
        public string HashAlgorithm { get; set; } = "SHA-256";
        public bool AlwaysSign { get; set; }
        public bool AlwaysEncrypt { get; set; }
        public CertValidationMode ValidationMode { get; set; } = CertValidationMode.Relaxed;

        /// <summary>
        /// When true, uses Outlook's native S/MIME engine (PR_SECURITY_FLAGS) for encryption.
        /// This produces standard S/MIME messages that any client (Entrust, native Outlook, etc.)
        /// can decrypt. When false, uses Parcl's own CMS envelope (smime.p7m attachment) which
        /// supports protected headers (RFC 7508) but requires Parcl on the receiving end.
        /// </summary>
        public bool UseNativeSmime { get; set; } = true;
    }

    public enum CertValidationMode
    {
        None,       // Expiry only — for self-signed certs, lab/test
        Relaxed,    // Chain validation, skip revocation — for internal CAs
        Strict      // Full chain + OCSP/CRL — for public PKI
    }

    public class CacheSettings
    {
        public bool EnableCertCache { get; set; } = true;
        public int CacheExpirationHours { get; set; } = 24;
        public int MaxCacheEntries { get; set; } = 500;
    }

    public class BehaviorSettings
    {
        public bool AutoDecrypt { get; set; }
        public string LogLevel { get; set; } = "Info";
        public LookupTrigger AutoLookup { get; set; } = LookupTrigger.OnCompose;
        public bool PromptOnMissingCert { get; set; } = true;
        public bool ShowStatusBar { get; set; } = true;
    }

    public enum LookupTrigger
    {
        Manual,
        OnCompose,
        OnSend
    }
}
