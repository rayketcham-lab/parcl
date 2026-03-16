using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
            try
            {
                _personalStore.Open(OpenFlags.ReadOnly);
                var results = _personalStore.Certificates.Find(
                    X509FindType.FindBySubjectName, email, false);

                return results.Cast<X509Certificate2>()
                    .Where(c => c.NotAfter > DateTime.UtcNow && c.NotBefore <= DateTime.UtcNow)
                    .OrderByDescending(c => c.NotAfter)
                    .FirstOrDefault();
            }
            finally
            {
                _personalStore.Close();
            }
        }

        public void ImportCertificate(byte[] certData)
        {
            var cert = new X509Certificate2(certData);
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

        public void Dispose()
        {
            _personalStore?.Dispose();
        }
    }
}
