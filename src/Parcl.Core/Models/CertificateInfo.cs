using System;
using System.Security.Cryptography.X509Certificates;

namespace Parcl.Core.Models
{
    public class CertificateInfo
    {
        public string Thumbprint { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Issuer { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public DateTime NotBefore { get; set; }
        public DateTime NotAfter { get; set; }
        public string SerialNumber { get; set; } = string.Empty;
        public X509KeyUsageFlags KeyUsage { get; set; }
        public bool HasPrivateKey { get; set; }
        /// <summary>
        /// Raw certificate bytes. Excluded from JSON serialization to avoid
        /// persisting sensitive cert data in cache files.
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        public byte[]? RawData { get; set; }

        public bool IsExpired => DateTime.UtcNow > NotAfter;
        public bool IsNotYetValid => DateTime.UtcNow < NotBefore;
        public bool IsValid => !IsExpired && !IsNotYetValid;

        public bool IsSigningCert =>
            KeyUsage.HasFlag(X509KeyUsageFlags.DigitalSignature) ||
            KeyUsage.HasFlag(X509KeyUsageFlags.NonRepudiation);

        public bool IsEncryptionCert =>
            KeyUsage.HasFlag(X509KeyUsageFlags.KeyEncipherment) ||
            KeyUsage.HasFlag(X509KeyUsageFlags.DataEncipherment);

        public static CertificateInfo FromX509(X509Certificate2 cert)
        {
            var keyUsage = X509KeyUsageFlags.None;
            foreach (var ext in cert.Extensions)
            {
                if (ext is X509KeyUsageExtension ku)
                {
                    keyUsage = ku.KeyUsages;
                    break;
                }
            }

            var email = string.Empty;
            foreach (var ext in cert.Extensions)
            {
                if (ext.Oid?.Value == "2.5.29.17") // Subject Alternative Name
                {
                    var san = ext.Format(false);
                    // Parse RFC822 Name from SAN string which may contain multiple entries, e.g.:
                    // "Other Name:Principal Name=E21127560@adxuser.com, RFC822 Name=james@rtx.com"
                    // Must find the "RFC822 Name=" prefix and extract its value specifically.
                    var rfc822Prefix = "RFC822 Name=";
                    var idx = san.IndexOf(rfc822Prefix, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0)
                    {
                        var value = san.Substring(idx + rfc822Prefix.Length);
                        // Trim at the next SAN entry separator (", " followed by a field name)
                        var commaIdx = value.IndexOf(',');
                        email = commaIdx >= 0 ? value.Substring(0, commaIdx).Trim() : value.Trim();
                    }
                }
            }

            return new CertificateInfo
            {
                Thumbprint = cert.Thumbprint,
                Subject = cert.Subject,
                Issuer = cert.Issuer,
                Email = email,
                NotBefore = cert.NotBefore,
                NotAfter = cert.NotAfter,
                SerialNumber = cert.SerialNumber,
                KeyUsage = keyUsage,
                HasPrivateKey = cert.HasPrivateKey,
                RawData = cert.RawData
            };
        }

        public override string ToString() =>
            $"{Subject} [{Thumbprint.Substring(0, 8)}...] Expires: {NotAfter:yyyy-MM-dd}";
    }
}
