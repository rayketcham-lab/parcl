using Parcl.Core.Config;

namespace Parcl.Core.Models
{
    public class LdapDirectoryEntry
    {
        public string Name { get; set; } = string.Empty;
        public string Server { get; set; } = string.Empty;
        public int Port { get; set; } = 636;
        public bool UseSsl { get; set; } = true;
        public string BaseDn { get; set; } = string.Empty;
        public string SearchFilter { get; set; } = "(mail={0})";
        public string CertAttribute { get; set; } = "userCertificate;binary";
        public AuthType AuthType { get; set; } = AuthType.Negotiate;
        public string? BindDn { get; set; }

        /// <summary>
        /// Stores the DPAPI-encrypted, base64-encoded bind password.
        /// Use SetBindPassword / GetBindPassword for plaintext access.
        /// </summary>
        public string? BindPassword { get; set; }

        public bool Enabled { get; set; } = true;

        public string ConnectionString =>
            $"{(UseSsl ? "ldaps" : "ldap")}://{Server}:{Port}";

        /// <summary>
        /// Encrypts and stores a plaintext password using DPAPI.
        /// </summary>
        public void SetBindPassword(string plaintext)
        {
            BindPassword = string.IsNullOrEmpty(plaintext)
                ? null
                : CredentialProtector.Protect(plaintext);
        }

        /// <summary>
        /// Decrypts and returns the stored bind password.
        /// </summary>
        public string GetBindPassword()
        {
            if (string.IsNullOrEmpty(BindPassword))
                return string.Empty;

            try
            {
                return CredentialProtector.Unprotect(BindPassword!);
            }
            catch
            {
                // If decryption fails (e.g., data was stored as plaintext before migration),
                // return empty to force re-entry.
                return string.Empty;
            }
        }
    }

    public enum AuthType
    {
        Anonymous,
        Simple,
        Negotiate
    }
}
