using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Parcl.Core.Config;
using Parcl.Core.Models;

namespace Parcl.Core.Crypto
{
    public class CertificateStore : IDisposable
    {
        private readonly X509Store _personalStore;

        public CertificateStore()
        {
            _personalStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        }

        public List<CertificateInfo> GetSigningCertificates()
        {
            return GetCertificates(cert => CertificateInfo.FromX509(cert).IsSigningCert);
        }

        public List<CertificateInfo> GetEncryptionCertificates()
        {
            return GetCertificates(cert => CertificateInfo.FromX509(cert).IsEncryptionCert);
        }

        public List<CertificateInfo> GetAllCertificates()
        {
            return GetCertificates(_ => true);
        }

        public X509Certificate2? FindByThumbprint(string thumbprint)
        {
            try
            {
                _personalStore.Open(OpenFlags.ReadOnly);
                var results = _personalStore.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, false);
                return results.Count > 0 ? results[0] : null;
            }
            finally
            {
                _personalStore.Close();
            }
        }

        public X509Certificate2? FindByEmail(string email)
        {
            // Search multiple stores: AddressBook (Other People) → My (Personal) → Root
            // This supports local users with certs, imported recipient certs, and enterprise PKI.
            var storeNames = new[] { StoreName.AddressBook, StoreName.My };
            foreach (var storeName in storeNames)
            {
                var cert = FindByEmailInStore(email, storeName);
                if (cert != null)
                    return cert;
            }
            return null;
        }

        private static X509Certificate2? FindByEmailInStore(string email, StoreName storeName)
        {
            // RFC 8398: Normalize internationalized email addresses to NFC
            // before comparison so that equivalent Unicode representations match.
            var normalizedEmail = email.Trim().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
            using (var store = new X509Store(storeName, StoreLocation.CurrentUser))
            {
                try
                {
                    store.Open(OpenFlags.ReadOnly);

                    // First try FindBySubjectName (matches CN in Subject DN)
                    var results = store.Certificates.Find(
                        X509FindType.FindBySubjectName, email, true);
                    var match = results.Cast<X509Certificate2>()
                        .Where(c => c.NotAfter > DateTime.UtcNow && c.NotBefore <= DateTime.UtcNow)
                        .OrderByDescending(c => c.NotAfter)
                        .FirstOrDefault();
                    if (match != null) return match;

                    // Fall back to scanning all certs and checking the SAN email field,
                    // since many certs store the email only in Subject Alternative Name.
                    foreach (var cert in store.Certificates.Cast<X509Certificate2>())
                    {
                        if (cert.NotAfter <= DateTime.UtcNow || cert.NotBefore > DateTime.UtcNow)
                            continue;

                        if (CertContainsEmail(cert, normalizedEmail))
                            return cert;
                    }

                    return null;
                }
                finally
                {
                    store.Close();
                }
            }
        }

        private static bool CertContainsEmail(X509Certificate2 cert, string email)
        {
            // Check Subject DN for email-like content (E= field or inline)
            if (cert.Subject != null &&
                cert.Subject.ToLowerInvariant().Contains(email))
                return true;

            // Check Subject Alternative Name extension for RFC822 email
            foreach (var ext in cert.Extensions)
            {
                if (ext.Oid?.Value == "2.5.29.17") // SAN
                {
                    var san = ext.Format(false).ToLowerInvariant();
                    if (san.Contains(email))
                        return true;
                }
            }

            // Check RDN components: extract CN from subject and match against
            // the local part of the email or the display name.
            // Handles enterprise certs where CN="James R Ketcham" but email
            // is user.name@quantumnexum.com (dot-separated name convention).
            if (cert.Subject != null && email.Contains("@"))
            {
                var localPart = email.Substring(0, email.IndexOf('@')).ToLowerInvariant();

                // Parse CN from subject DN (e.g., "CN=James R Ketcham + uid=...")
                var subject = cert.Subject;
                int cnStart = subject.IndexOf("CN=", StringComparison.OrdinalIgnoreCase);
                if (cnStart >= 0)
                {
                    cnStart += 3;
                    int cnEnd = subject.IndexOfAny(new[] { ',', '+' }, cnStart);
                    if (cnEnd < 0) cnEnd = subject.Length;
                    var cn = subject.Substring(cnStart, cnEnd - cnStart).Trim().ToLowerInvariant();

                    // Match: "james.r.ketcham" vs "james r ketcham" (dots = spaces)
                    var cnNormalized = cn.Replace(" ", ".").Replace("-", ".");
                    if (cnNormalized == localPart || cn.Replace(" ", "") == localPart.Replace(".", ""))
                        return true;
                }

                // Also check E= field in subject DN
                int eStart = subject.IndexOf("E=", StringComparison.OrdinalIgnoreCase);
                if (eStart >= 0)
                {
                    eStart += 2;
                    int eEnd = subject.IndexOfAny(new[] { ',', '+', ' ' }, eStart);
                    if (eEnd < 0) eEnd = subject.Length;
                    var emailInDn = subject.Substring(eStart, eEnd - eStart).Trim().ToLowerInvariant();
                    if (emailInDn == email)
                        return true;
                }
            }

            return false;
        }

        public void ImportCertificate(byte[] certData, string? password = null)
        {
            var cert = password != null
                ? new X509Certificate2(certData, password,
                    X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.PersistKeySet)
                : new X509Certificate2(certData,
                    (string?)null,
                    X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.PersistKeySet);
            try
            {
                _personalStore.Open(OpenFlags.ReadWrite);
                _personalStore.Add(cert);
            }
            finally
            {
                _personalStore.Close();
            }
        }

        public byte[]? ExportPublicCertificate(string thumbprint)
        {
            var cert = FindByThumbprint(thumbprint);
            return cert?.Export(X509ContentType.Cert);
        }

        /// <summary>
        /// Publishes a certificate to the "Other People" (AddressBook) store so
        /// Outlook's S/MIME engine can find it when encrypting to this recipient.
        /// </summary>
        public void PublishToAddressBook(X509Certificate2 cert)
        {
            using (var addressBook = new X509Store(StoreName.AddressBook, StoreLocation.CurrentUser))
            {
                try
                {
                    addressBook.Open(OpenFlags.ReadWrite);
                    var existing = addressBook.Certificates.Find(
                        X509FindType.FindByThumbprint, cert.Thumbprint, false);
                    if (existing.Count == 0)
                        addressBook.Add(cert);
                }
                finally
                {
                    addressBook.Close();
                }
            }
        }

        private List<CertificateInfo> GetCertificates(Func<X509Certificate2, bool> filter)
        {
            var results = new List<CertificateInfo>();
            try
            {
                _personalStore.Open(OpenFlags.ReadOnly);
                foreach (var cert in _personalStore.Certificates.Cast<X509Certificate2>())
                {
                    if (!cert.NotAfter.Equals(default) && cert.NotAfter > DateTime.UtcNow && filter(cert))
                    {
                        results.Add(CertificateInfo.FromX509(cert));
                    }
                }
            }
            finally
            {
                _personalStore.Close();
            }

            return results.OrderByDescending(c => c.NotAfter).ToList();
        }

        /// <summary>
        /// Builds an X.509 chain and returns the result.
        /// Validation depth is controlled by <see cref="CertValidationMode"/>:
        /// None = expiry only, Relaxed = chain without revocation, Strict = chain + OCSP/CRL.
        /// </summary>
        public ChainValidationResult ValidateCertificateChain(X509Certificate2 cert, CertValidationMode mode)
        {
            if (mode == CertValidationMode.None)
            {
                bool expired = cert.NotAfter <= DateTime.UtcNow || cert.NotBefore > DateTime.UtcNow;
                return expired
                    ? new ChainValidationResult { IsValid = false, ErrorMessage = "Certificate is expired or not yet valid." }
                    : new ChainValidationResult { IsValid = true };
            }

            using (var chain = new X509Chain())
            {
                if (mode == CertValidationMode.Relaxed)
                {
                    chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                }
                else // Strict
                {
                    chain.ChainPolicy.RevocationMode = X509RevocationMode.Online;
                    chain.ChainPolicy.RevocationFlag = X509RevocationFlag.EntireChain;
                }

                chain.ChainPolicy.VerificationFlags = X509VerificationFlags.NoFlag;

                bool isValid = chain.Build(cert);
                if (isValid)
                    return new ChainValidationResult { IsValid = true };

                var errors = new System.Text.StringBuilder();
                foreach (var status in chain.ChainStatus)
                {
                    errors.AppendLine(status.StatusInformation);
                }

                return new ChainValidationResult
                {
                    IsValid = false,
                    ErrorMessage = errors.ToString().TrimEnd()
                };
            }
        }

        public void Dispose()
        {
            _personalStore?.Dispose();
        }
    }

    public class ChainValidationResult
    {
        public bool IsValid { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
