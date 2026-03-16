using System;
using System.IO;
using System.Security.Cryptography;

namespace Parcl.Core.Config
{
    /// <summary>
    /// Provides HMAC-SHA256 integrity protection for the settings file.
    /// Uses a DPAPI-protected random key stored in %APPDATA%\Parcl\settings.key.
    /// The HMAC is written to a companion file: settings.json.hmac (hex-encoded).
    /// </summary>
    internal static class SettingsIntegrity
    {
        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl");

        private static readonly string KeyFile = Path.Combine(SettingsDir, "settings.key");
        private static readonly string HmacFile = Path.Combine(SettingsDir, "settings.json.hmac");

        private static readonly byte[] DpapiEntropy =
            System.Text.Encoding.UTF8.GetBytes("Parcl.Core.Config.SettingsIntegrity");

        /// <summary>
        /// Computes HMAC-SHA256 of the given settings JSON and writes it to the .hmac file.
        /// Creates the HMAC key on first use.
        /// </summary>
        public static void WriteHmac(string settingsJson)
        {
            try
            {
                var key = GetOrCreateKey();
                var hmac = ComputeHmac(key, settingsJson);
                var hex = BitConverter.ToString(hmac).Replace("-", "").ToLowerInvariant();

                Directory.CreateDirectory(SettingsDir);
                File.WriteAllText(HmacFile, hex);
            }
            catch (Exception)
            {
                // Non-fatal: if we can't write the HMAC, settings still work.
                // This can happen on first-run race conditions or permission issues.
            }
        }

        /// <summary>
        /// Verifies the HMAC of the given settings JSON against the stored .hmac file.
        /// Returns true if valid, false if tampered or if HMAC file is missing/corrupt.
        /// </summary>
        public static bool VerifyHmac(string settingsJson, out string failureReason)
        {
            failureReason = string.Empty;

            try
            {
                if (!File.Exists(HmacFile))
                {
                    failureReason = "HMAC file not found — settings predates integrity protection";
                    return false;
                }

                if (!File.Exists(KeyFile))
                {
                    failureReason = "HMAC key file not found";
                    return false;
                }

                var key = LoadKey();
                if (key == null)
                {
                    failureReason = "Failed to unprotect HMAC key";
                    return false;
                }

                var expectedHmac = ComputeHmac(key, settingsJson);
                var expectedHex = BitConverter.ToString(expectedHmac).Replace("-", "").ToLowerInvariant();

                var storedHex = File.ReadAllText(HmacFile).Trim().ToLowerInvariant();

                if (!ConstantTimeEquals(expectedHex, storedHex))
                {
                    failureReason = "Settings file may have been tampered with";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                failureReason = $"HMAC verification error: {ex.GetType().Name}";
                return false;
            }
        }

        private static byte[] GetOrCreateKey()
        {
            if (File.Exists(KeyFile))
            {
                var existing = LoadKey();
                if (existing != null)
                    return existing;
            }

            // Generate a new 32-byte random key
            var keyBytes = new byte[32];
            using (var rng = RandomNumberGenerator.Create())
                rng.GetBytes(keyBytes);

            // Protect with DPAPI (CurrentUser scope) and write to disk
            var protectedKey = ProtectedData.Protect(keyBytes, DpapiEntropy, DataProtectionScope.CurrentUser);
            Directory.CreateDirectory(SettingsDir);
            File.WriteAllBytes(KeyFile, protectedKey);

            return keyBytes;
        }

        private static byte[]? LoadKey()
        {
            try
            {
                var protectedKey = File.ReadAllBytes(KeyFile);
                return ProtectedData.Unprotect(protectedKey, DpapiEntropy, DataProtectionScope.CurrentUser);
            }
            catch
            {
                return null;
            }
        }

        private static byte[] ComputeHmac(byte[] key, string content)
        {
            using (var hmac = new HMACSHA256(key))
            {
                var contentBytes = System.Text.Encoding.UTF8.GetBytes(content);
                return hmac.ComputeHash(contentBytes);
            }
        }

        /// <summary>
        /// Constant-time string comparison to prevent timing attacks on HMAC values.
        /// </summary>
        private static bool ConstantTimeEquals(string a, string b)
        {
            if (a.Length != b.Length)
                return false;

            int diff = 0;
            for (int i = 0; i < a.Length; i++)
                diff |= a[i] ^ b[i];

            return diff == 0;
        }
    }
}
