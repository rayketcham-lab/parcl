using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Parcl.Core.Tests
{
    /// <summary>
    /// Shared test utilities for certificate creation and store management.
    /// Eliminates duplication across test classes.
    /// </summary>
    public static class TestCertFactory
    {
        public static X509Certificate2 Create(
            string subject,
            X509KeyUsageFlags keyUsage = X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
            DateTimeOffset? notBefore = null,
            DateTimeOffset? notAfter = null)
        {
            using (var rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(subject, rsa,
                    HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                request.CertificateExtensions.Add(
                    new X509KeyUsageExtension(keyUsage, critical: true));

                var cert = request.CreateSelfSigned(
                    notBefore ?? DateTimeOffset.UtcNow.AddMinutes(-5),
                    notAfter ?? DateTimeOffset.UtcNow.AddHours(1));

                return new X509Certificate2(
                    cert.Export(X509ContentType.Pfx, "test"), "test",
                    X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);
            }
        }
    }

    public static class TestCertStore
    {
        public static void Add(params X509Certificate2[] certs)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                foreach (var cert in certs)
                    store.Add(cert);
            }
        }

        public static void Remove(params X509Certificate2[] certs)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                foreach (var cert in certs)
                {
                    var matches = store.Certificates.Find(
                        X509FindType.FindByThumbprint, cert.Thumbprint, false);
                    foreach (X509Certificate2 c in matches)
                        store.Remove(c);
                }
            }
        }

        public static void RemoveByThumbprint(string thumbprint)
        {
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                var matches = store.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, false);
                foreach (var c in matches.Cast<X509Certificate2>())
                    store.Remove(c);
            }
        }
    }
}
