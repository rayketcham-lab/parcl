using System;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Parcl.Core.Models;

namespace Parcl.Core.Crypto
{
    public class CertExchange
    {
        private readonly CertificateStore _certStore;

        public CertExchange(CertificateStore certStore)
        {
            _certStore = certStore;
        }

        public CertExchangePayload PrepareExport(string thumbprint)
        {
            var certData = _certStore.ExportPublicCertificate(thumbprint);
            if (certData == null)
                throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found.");

            var cert = _certStore.FindByThumbprint(thumbprint);
            var info = CertificateInfo.FromX509(cert!);

            return new CertExchangePayload
            {
                CertificateData = Convert.ToBase64String(certData),
                SenderEmail = info.Email,
                SenderName = info.Subject,
                Thumbprint = info.Thumbprint,
                ExpirationDate = info.NotAfter
            };
        }

        public CertificateInfo ImportFromPayload(CertExchangePayload payload)
        {
            var certData = Convert.FromBase64String(payload.CertificateData);
            _certStore.ImportCertificate(certData);

            var cert = new X509Certificate2(certData);
            return CertificateInfo.FromX509(cert);
        }

        public string FormatAsAttachment(CertExchangePayload payload)
        {
            var sb = new StringBuilder();
            sb.AppendLine("-----BEGIN CERTIFICATE-----");
            // Wrap base64 at 64 chars per line
            var b64 = payload.CertificateData;
            for (int i = 0; i < b64.Length; i += 64)
            {
                sb.AppendLine(b64.Substring(i, Math.Min(64, b64.Length - i)));
            }
            sb.AppendLine("-----END CERTIFICATE-----");
            return sb.ToString();
        }
    }

    public class CertExchangePayload
    {
        public string CertificateData { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string Thumbprint { get; set; } = string.Empty;
        public DateTime ExpirationDate { get; set; }
    }
}
