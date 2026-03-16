using System;
using System.Security.Cryptography;
using System.Text;

namespace Parcl.Core.Config
{
    /// <summary>
    /// Encrypts and decrypts credentials using DPAPI (CurrentUser scope).
    /// Values are stored as base64-encoded ciphertext.
    /// </summary>
    public static class CredentialProtector
    {
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("Parcl.Core.Config.CredentialProtector");

        public static string Protect(string plaintext)
        {
            if (string.IsNullOrEmpty(plaintext))
                return string.Empty;

            var plaintextBytes = Encoding.UTF8.GetBytes(plaintext);
            var cipherBytes = ProtectedData.Protect(
                plaintextBytes, Entropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(cipherBytes);
        }

        public static string Unprotect(string protectedBase64)
        {
            if (string.IsNullOrEmpty(protectedBase64))
                return string.Empty;

            var cipherBytes = Convert.FromBase64String(protectedBase64);
            var plaintextBytes = ProtectedData.Unprotect(
                cipherBytes, Entropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plaintextBytes);
        }
    }
}
