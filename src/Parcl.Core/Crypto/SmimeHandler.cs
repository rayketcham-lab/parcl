using System;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace Parcl.Core.Crypto
{
    public class SmimeHandler
    {
        public byte[] Sign(byte[] content, X509Certificate2 signingCert)
        {
            if (!signingCert.HasPrivateKey)
                throw new InvalidOperationException("Signing certificate must have a private key.");

            var contentInfo = new ContentInfo(content);
            var signedCms = new SignedCms(contentInfo, detached: false);
            var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, signingCert)
            {
                DigestAlgorithm = new Oid("2.16.840.1.101.3.4.2.1") // SHA-256
            };
            signer.IncludeOption = X509IncludeOption.WholeChain;

            signedCms.ComputeSignature(signer);
            return signedCms.Encode();
        }

        public SmimeVerifyResult Verify(byte[] signedData)
        {
            try
            {
                var signedCms = new SignedCms();
                signedCms.Decode(signedData);
                signedCms.CheckSignature(verifySignatureOnly: false);

                if (signedCms.SignerInfos.Count == 0)
                {
                    return new SmimeVerifyResult
                    {
                        IsValid = false,
                        ErrorMessage = "Signed message contains no signer information."
                    };
                }

                var signerCert = signedCms.SignerInfos[0].Certificate;
                return new SmimeVerifyResult
                {
                    IsValid = true,
                    Content = signedCms.ContentInfo.Content,
                    SignerCertificate = signerCert != null
                        ? Models.CertificateInfo.FromX509(signerCert)
                        : null
                };
            }
            catch (CryptographicException ex)
            {
                return new SmimeVerifyResult
                {
                    IsValid = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        public byte[] Encrypt(byte[] content, X509Certificate2Collection recipientCerts)
        {
            if (recipientCerts.Count == 0)
                throw new ArgumentException("At least one recipient certificate is required.");

            // Validate each recipient certificate chain before encrypting
            using (var certStore = new CertificateStore())
            {
                foreach (X509Certificate2 cert in recipientCerts)
                {
                    var chainResult = certStore.ValidateCertificateChain(cert);
                    if (!chainResult.IsValid)
                    {
                        throw new CryptographicException(
                            $"Certificate chain validation failed for '{cert.Subject}': {chainResult.ErrorMessage}");
                    }
                }
            }

            var contentInfo = new ContentInfo(content);
            var envelopedCms = new EnvelopedCms(contentInfo,
                new AlgorithmIdentifier(new Oid("2.16.840.1.101.3.4.1.42"))); // AES-256-CBC

            var recipients = new CmsRecipientCollection();
            foreach (X509Certificate2 cert in recipientCerts)
            {
                recipients.Add(new CmsRecipient(SubjectIdentifierType.IssuerAndSerialNumber, cert));
            }

            envelopedCms.Encrypt(recipients);
            return envelopedCms.Encode();
        }

        public SmimeDecryptResult Decrypt(byte[] encryptedData)
        {
            try
            {
                var envelopedCms = new EnvelopedCms();
                envelopedCms.Decode(encryptedData);

                // Decrypt uses the current user's certificate store automatically
                envelopedCms.Decrypt();
                return new SmimeDecryptResult
                {
                    Success = true,
                    Content = envelopedCms.ContentInfo.Content
                };
            }
            catch (CryptographicException ex)
            {
                return new SmimeDecryptResult
                {
                    Success = false,
                    ErrorMessage = $"Decryption failed: {ex.Message}"
                };
            }
        }
    }

    public class SmimeVerifyResult
    {
        public bool IsValid { get; set; }
        public byte[]? Content { get; set; }
        public Models.CertificateInfo? SignerCertificate { get; set; }
        public string? ErrorMessage { get; set; }
    }

    public class SmimeDecryptResult
    {
        public bool Success { get; set; }
        public byte[]? Content { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
