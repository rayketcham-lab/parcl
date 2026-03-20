using System;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace Parcl.Core.Crypto
{
    public class SmimeHandler
    {
        private readonly string _encryptionAlgorithm;
        private readonly string _hashAlgorithm;

        // OID lookup for encryption algorithms (secure only)
        private static readonly System.Collections.Generic.Dictionary<string, string> EncryptionOids =
            new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["AES-128-CBC"] = "2.16.840.1.101.3.4.1.2",
                ["AES-192-CBC"] = "2.16.840.1.101.3.4.1.22",
                ["AES-256-CBC"] = "2.16.840.1.101.3.4.1.42",
                // AES-GCM authenticated encryption (NIST SP 800-38D).
                // Note: .NET Framework 4.8 EnvelopedCms may not support GCM OIDs at runtime.
                // If the runtime does not support them, the Encrypt() call will throw a
                // CryptographicException and the user will see an error message.
                ["AES-128-GCM"] = "2.16.840.1.101.3.4.1.6",
                ["AES-256-GCM"] = "2.16.840.1.101.3.4.1.46",
            };

        // OID lookup for hash algorithms (secure only — SHA-2 family)
        private static readonly System.Collections.Generic.Dictionary<string, string> HashOids =
            new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["SHA-256"] = "2.16.840.1.101.3.4.2.1",
                ["SHA-384"] = "2.16.840.1.101.3.4.2.2",
                ["SHA-512"] = "2.16.840.1.101.3.4.2.3",
            };

        // Blocklisted insecure hash algorithms — these MUST never be used for signing
        private static readonly System.Collections.Generic.HashSet<string> BlockedHashAlgorithms =
            new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "MD2", "MD4", "MD5", "SHA-1", "SHA1",
                // OIDs for blocked hash algorithms
                "1.2.840.113549.2.2",   // MD2
                "1.2.840.113549.2.4",   // MD4
                "1.2.840.113549.2.5",   // MD5
                "1.3.14.3.2.26",        // SHA-1
            };

        // Blocklisted insecure encryption algorithms — these MUST never be used
        private static readonly System.Collections.Generic.HashSet<string> BlockedEncryptionAlgorithms =
            new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "DES", "DES-CBC", "3DES", "3DES-CBC", "RC2", "RC2-CBC", "RC4",
                // OIDs for blocked encryption algorithms
                "1.3.14.3.2.7",         // DES-CBC
                "1.2.840.113549.3.7",   // 3DES-CBC
                "1.2.840.113549.3.2",   // RC2-CBC
                "1.2.840.113549.3.4",   // RC4
            };

        public SmimeHandler(string encryptionAlgorithm = "AES-256-CBC", string hashAlgorithm = "SHA-256")
        {
            // Reject insecure algorithms at construction time
            if (BlockedHashAlgorithms.Contains(hashAlgorithm))
                throw new ArgumentException(
                    $"Hash algorithm '{hashAlgorithm}' is blocked — it is cryptographically insecure. " +
                    "Use SHA-256, SHA-384, or SHA-512.", nameof(hashAlgorithm));

            if (BlockedEncryptionAlgorithms.Contains(encryptionAlgorithm))
                throw new ArgumentException(
                    $"Encryption algorithm '{encryptionAlgorithm}' is blocked — it is cryptographically insecure. " +
                    "Use AES-128-CBC, AES-192-CBC, AES-256-CBC, AES-128-GCM, or AES-256-GCM.", nameof(encryptionAlgorithm));

            _encryptionAlgorithm = encryptionAlgorithm;
            _hashAlgorithm = hashAlgorithm;
        }

        /// <summary>
        /// Returns true if the given hash algorithm name or OID is blocked as insecure.
        /// </summary>
        public static bool IsBlockedHashAlgorithm(string algorithmNameOrOid) =>
            BlockedHashAlgorithms.Contains(algorithmNameOrOid);

        /// <summary>
        /// Returns true if the given encryption algorithm name or OID is blocked as insecure.
        /// </summary>
        public static bool IsBlockedEncryptionAlgorithm(string algorithmNameOrOid) =>
            BlockedEncryptionAlgorithms.Contains(algorithmNameOrOid);

        public byte[] Sign(byte[] content, X509Certificate2 signingCert)
        {
            if (!signingCert.HasPrivateKey)
                throw new InvalidOperationException("Signing certificate must have a private key.");

            var hashOid = HashOids.TryGetValue(_hashAlgorithm, out var h) ? h : "2.16.840.1.101.3.4.2.1";

            var contentInfo = new ContentInfo(content);
            var signedCms = new SignedCms(contentInfo, detached: false);
            var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, signingCert)
            {
                DigestAlgorithm = new Oid(hashOid)
            };
            signer.IncludeOption = X509IncludeOption.WholeChain;

            // RFC 6211: Algorithm Protection attribute (OID 1.2.840.113549.1.9.52)
            // This signed attribute binds the digest and signature algorithms to the
            // signature, preventing algorithm substitution attacks.
            // .NET Framework 4.8 CmsSigner.SignedAttributes is available, but building
            // the correct ASN.1 structure for the CMSAlgorithmProtection attribute
            // requires manual DER encoding. The attribute value contains:
            //   digestAlgorithm AlgorithmIdentifier, signatureAlgorithm [1] AlgorithmIdentifier
            // TODO: Implement full ASN.1 DER encoding for CMSAlgorithmProtection when
            // a suitable ASN.1 library is available. .NET 4.8 does not provide a public
            // ASN.1 writer, and hand-rolling DER is error-prone for production use.

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

                // Check if the signature used an insecure hash algorithm
                var digestOid = signedCms.SignerInfos[0].DigestAlgorithm.Value;
                var digestName = signedCms.SignerInfos[0].DigestAlgorithm.FriendlyName ?? digestOid;
                bool usedInsecureHash = IsBlockedHashAlgorithm(digestOid ?? "") ||
                                        IsBlockedHashAlgorithm(digestName ?? "");

                var signerCert = signedCms.SignerInfos[0].Certificate;
                return new SmimeVerifyResult
                {
                    IsValid = true,
                    UsedInsecureAlgorithm = usedInsecureHash,
                    DigestAlgorithm = digestName,
                    Content = signedCms.ContentInfo.Content,
                    SignerCertificate = signerCert != null
                        ? Models.CertificateInfo.FromX509(signerCert)
                        : null,
                    InsecureAlgorithmWarning = usedInsecureHash
                        ? $"WARNING: This message was signed with {digestName}, which is cryptographically insecure. " +
                          "Ask the sender to re-sign using SHA-256 or stronger."
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

            var encOid = EncryptionOids.TryGetValue(_encryptionAlgorithm, out var e) ? e : "2.16.840.1.101.3.4.1.42";

            var contentInfo = new ContentInfo(content);
            var envelopedCms = new EnvelopedCms(contentInfo,
                new AlgorithmIdentifier(new Oid(encOid)));

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
        /// <summary>True if the signature used a blocked/insecure hash algorithm (e.g. SHA-1, MD5).</summary>
        public bool UsedInsecureAlgorithm { get; set; }
        /// <summary>The digest algorithm used by the signer (e.g. "sha256", "sha1").</summary>
        public string? DigestAlgorithm { get; set; }
        /// <summary>Warning message if an insecure algorithm was detected.</summary>
        public string? InsecureAlgorithmWarning { get; set; }
    }

    public class SmimeDecryptResult
    {
        public bool Success { get; set; }
        public byte[]? Content { get; set; }
        public string? ErrorMessage { get; set; }
    }
}
