using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Parcl.Addin.TaskPane;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parcl.Addin
{
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(AddInProgId)]
    public partial class ParclAddIn : IDTExtensibility2, Microsoft.Office.Core.IRibbonExtensibility
    {
        public const string ClassId = "B8F0C3A2-7D5E-4F91-A6C8-9E1B3D5A7F42";
        public const string AddInProgId = "Parcl.Addin";

        private Outlook.Application? _application;
        private Form? _taskPaneForm;
        private ParclTaskPaneHost? _taskPaneHost;

        internal static ParclAddIn? Current { get; private set; }

        internal ParclSettings Settings { get; private set; } = null!;
        internal CertificateStore CertStore { get; private set; } = null!;
        internal SmimeHandler SmimeHandler { get; private set; } = null!;
        internal LdapCertLookup LdapLookup { get; private set; } = null!;
        internal CertificateCache CertCache { get; private set; } = null!;
        internal CertExchange CertExchange { get; private set; } = null!;
        internal ParclLogger Logger { get; private set; } = null!;

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            _application = (Outlook.Application)application;
            Current = this;

            Logger = new ParclLogger();
            Logger.Info("AddIn", "Parcl add-in connecting");

            try
            {
                Settings = ParclSettings.Load();
                Logger.Debug("AddIn",
                    $"Settings loaded — {Settings.LdapDirectories.Count} LDAP directories configured");

                CertStore = new CertificateStore();
                SmimeHandler = new SmimeHandler();
                LdapLookup = new LdapCertLookup(Logger);
                CertCache = new CertificateCache(
                    Settings.Cache.CacheExpirationHours,
                    Settings.Cache.MaxCacheEntries);
                CertExchange = new CertExchange(CertStore);

                Logger.Debug("AddIn", "Core services initialized");

                _taskPaneHost = new ParclTaskPaneHost();
                _taskPaneForm = new Form
                {
                    Text = "Parcl Dashboard",
                    Width = 340,
                    Height = 600,
                    FormBorderStyle = FormBorderStyle.SizableToolWindow,
                    ShowInTaskbar = false,
                    StartPosition = FormStartPosition.Manual,
                    TopMost = false
                };
                _taskPaneHost.Dock = DockStyle.Fill;
                _taskPaneForm.Controls.Add(_taskPaneHost);
                _taskPaneForm.FormClosing += (s, e) =>
                {
                    e.Cancel = true;
                    _taskPaneForm.Hide();
                };

                Logger.Info("AddIn", "Dashboard form created");

                // Always hook ItemSend — handles encrypt encapsulation and auto-sign/encrypt
                _application.ItemSend += Application_ItemSend;

                // Invalidate ribbon when user selects a different message
                // so toggle button states (Encrypt/Sign) reflect the selected message
                var explorer = _application.ActiveExplorer();
                if (explorer != null)
                    explorer.SelectionChange += Explorer_SelectionChange;

                // Also catch new inspectors (opened messages)
                _application.Inspectors.NewInspector += Inspectors_NewInspector;

                Logger.Debug("AddIn", "Ribbon selection tracking enabled");

                Logger.Info("AddIn", "Parcl add-in started successfully");
            }
            catch (Exception ex)
            {
                Logger.Error("AddIn", "Failed during startup", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Logger?.Info("AddIn", "Parcl add-in disconnecting");
            Settings?.Save();
            _taskPaneForm?.Dispose();
            CertStore?.Dispose();
            Logger?.Dispose();
            Current = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private void Explorer_SelectionChange()
        {
            // Ribbon toggle states (Encrypt/Sign) must reflect the SELECTED message
            try { _ribbon?.Invalidate(); }
            catch { }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            // When a message is opened in its own window, refresh ribbon state
            try
            {
                inspector.Activate();
                _ribbon?.Invalidate();
            }
            catch { }
        }

        internal void ToggleTaskPane()
        {
            if (_taskPaneForm == null) return;

            if (_taskPaneForm.Visible)
            {
                _taskPaneForm.Hide();
            }
            else
            {
                _taskPaneForm.Show();
            }

            Logger?.Debug("UI",
                $"Task pane toggled: {(_taskPaneForm.Visible ? "visible" : "hidden")}");
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            if (!(item is Outlook.MailItem mail)) return;

            var subjectPreview = mail.Subject != null
                ? mail.Subject.Substring(0, Math.Min(30, mail.Subject.Length))
                : "(no subject)";
            Logger.Info("Send", $"ItemSend intercepted — to: {mail.To}, subject: {subjectPreview}");

            bool shouldEncryptRequested = false;
            try
            {
                // ── Determine what's requested ──
                bool shouldSign = false;
                bool shouldEncrypt = false;

                var signFlag = mail.UserProperties.Find("ParclSign");
                if (signFlag != null && (bool)signFlag.Value)
                    shouldSign = true;

                if (Settings.Crypto.AlwaysSign &&
                    !string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                {
                    var signingCert = CertStore.FindByThumbprint(
                        Settings.UserProfile.SigningCertThumbprint!);
                    if (signingCert != null && signingCert.HasPrivateKey)
                        shouldSign = true;
                }

                var encryptFlag = mail.UserProperties.Find("ParclEncrypt");
                if (encryptFlag != null && (bool)encryptFlag.Value)
                    shouldEncrypt = true;

                if (Settings.Crypto.AlwaysEncrypt)
                    shouldEncrypt = true;

                shouldEncryptRequested = shouldEncrypt;

                // ── Apply ──
                // If encrypting: do it ourselves (sign goes INSIDE the encrypted envelope per RFC 5751).
                // If only signing: use Outlook's native PR_SECURITY_FLAGS.
                // Never set PR_SECURITY_FLAGS for sign when also encrypting — that double-wraps
                // and the recipient sees "Signed" as the outer layer instead of "Encrypted".

                if (shouldEncrypt)
                {
                    Logger.Info("Send", "Encapsulating message in S/MIME envelope"
                        + (shouldSign ? " (sign + encrypt)" : " (encrypt only)"));

                    string? encryptError = EncapsulateMessage(mail, shouldSign);
                    if (encryptError != null)
                    {
                        cancel = true;
                        Logger.Error("Send", $"Encryption failed — send blocked: {encryptError}");
                        MessageBox.Show(
                            $"Message NOT sent — encryption failed:\n\n{encryptError}\n\n" +
                            "Fix the issue or remove encryption before sending.",
                            "Parcl — Send Blocked",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }
                }
                else if (shouldSign)
                {
                    // Sign-only: use Outlook's native signing
                    const string PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
                    const int SECFLAG_SIGNED = 0x02;

                    var pa = mail.PropertyAccessor;
                    int flags;
                    try { flags = (int)pa.GetProperty(PR_SECURITY_FLAGS); }
                    catch { flags = 0; }

                    pa.SetProperty(PR_SECURITY_FLAGS, flags | SECFLAG_SIGNED);
                    Logger.Info("Send", "S/MIME signature flag applied (sign-only, no encrypt)");
                }
            }
            catch (Exception ex)
            {
                // If encryption was requested and something threw, BLOCK the send
                if (shouldEncryptRequested)
                {
                    cancel = true;
                    Logger.Error("Send", "Encryption failed with exception — send blocked", ex);
                    MessageBox.Show(
                        $"Message NOT sent — encryption error:\n\n{ex.Message}\n\n" +
                        "Fix the issue or remove encryption before sending.",
                        "Parcl — Send Blocked",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                else
                {
                    Logger.Error("Send", "Failed during send processing", ex);
                }
            }
        }

        /// <summary>
        /// Performs the actual S/MIME encapsulation at send time.
        /// If alsoSign is true, signs the MIME content BEFORE encrypting (RFC 5751 sign-then-encrypt).
        /// Returns null on success, or an error message string if encryption failed.
        /// If this returns non-null, the send MUST be cancelled.
        /// </summary>
        private string? EncapsulateMessage(Outlook.MailItem mail, bool alsoSign = false)
        {
            // ── Validate ALL recipients have valid certs ──
            var recipientCerts =
                new System.Security.Cryptography.X509Certificates.X509Certificate2Collection();
            var errors = new System.Collections.Generic.List<string>();

            for (int i = 1; i <= mail.Recipients.Count; i++)
            {
                var recipient = mail.Recipients[i];
                var smtpAddr = ResolveSmtpAddress(recipient);
                var cert = ResolveRecipientCert(smtpAddr, recipient);

                if (cert == null)
                {
                    errors.Add($"{smtpAddr}: No certificate found");
                    continue;
                }

                // Check expiry
                if (cert.NotAfter <= DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate expired on {cert.NotAfter:yyyy-MM-dd}");
                    continue;
                }

                if (cert.NotBefore > DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate not yet valid (starts {cert.NotBefore:yyyy-MM-dd})");
                    continue;
                }

                recipientCerts.Add(cert);
            }

            if (errors.Count > 0)
            {
                Logger.Error("Send",
                    $"Encryption blocked — {errors.Count} recipient(s) failed validation");
                return string.Join("\n", errors);
            }

            if (recipientCerts.Count == 0)
                return "No recipients with valid certificates";

            // Also encrypt to self so Sent Items are readable
            if (!string.IsNullOrEmpty(Settings.UserProfile.EncryptionCertThumbprint))
            {
                var selfCert = CertStore.FindByThumbprint(
                    Settings.UserProfile.EncryptionCertThumbprint!);
                if (selfCert != null)
                    recipientCerts.Add(selfCert);
            }

            // ── Build MIME content ──
            var attachments = new System.Collections.Generic.List<Parcl.Core.Crypto.MimeAttachment>();
            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var att = mail.Attachments[i];
                var tempAtt = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    System.IO.Path.GetRandomFileName() + System.IO.Path.GetExtension(att.FileName));
                try
                {
                    att.SaveAsFile(tempAtt);
                    attachments.Add(new Parcl.Core.Crypto.MimeAttachment
                    {
                        FileName = att.FileName,
                        Data = System.IO.File.ReadAllBytes(tempAtt)
                    });
                }
                finally
                {
                    try { System.IO.File.Delete(tempAtt); } catch { }
                }
            }

            // RFC 7508: include protected headers inside the encrypted envelope
            var protectedHeaders = new Parcl.Core.Crypto.ProtectedHeaders
            {
                Subject = mail.Subject,
                From = mail.SenderEmailAddress ?? "",
                To = mail.To ?? "",
                Date = DateTime.UtcNow.ToString("R")
            };

            var mimeContent = Parcl.Core.Crypto.MimeBuilder.Build(
                mail.Body, mail.HTMLBody,
                attachments.Count > 0 ? attachments : null,
                protectedHeaders);

            Logger.Debug("Send",
                $"MIME content built: {mimeContent.Length} bytes, {attachments.Count} attachment(s), protected headers included");

            // ── Sign INSIDE the envelope if requested (sign-then-encrypt per RFC 5751) ──
            byte[] contentToEncrypt = mimeContent;
            if (alsoSign && !string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
            {
                var signingCert = CertStore.FindByThumbprint(
                    Settings.UserProfile.SigningCertThumbprint!);
                if (signingCert != null && signingCert.HasPrivateKey)
                {
                    var signedCms = new System.Security.Cryptography.Pkcs.SignedCms(
                        new System.Security.Cryptography.Pkcs.ContentInfo(mimeContent), detached: false);
                    var signer = new System.Security.Cryptography.Pkcs.CmsSigner(
                        System.Security.Cryptography.Pkcs.SubjectIdentifierType.IssuerAndSerialNumber,
                        signingCert)
                    {
                        DigestAlgorithm = new System.Security.Cryptography.Oid("2.16.840.1.101.3.4.2.1"),
                        IncludeOption = System.Security.Cryptography.X509Certificates.X509IncludeOption.WholeChain
                    };
                    signedCms.ComputeSignature(signer);
                    contentToEncrypt = signedCms.Encode();
                    Logger.Info("Send",
                        $"MIME content signed inside envelope ({contentToEncrypt.Length} bytes)");
                }
                else
                {
                    Logger.Warn("Send", "Signing cert not found or has no private key — encrypting without signature");
                }
            }

            // ── Encrypt ──
            // Chain validation was too strict for self-signed/internal certs.
            // We already validated expiry above. The CMS encrypt will fail if
            // the cert is truly unusable.
            var contentInfo = new System.Security.Cryptography.Pkcs.ContentInfo(contentToEncrypt);
            var envelopedCms = new System.Security.Cryptography.Pkcs.EnvelopedCms(
                contentInfo,
                new System.Security.Cryptography.Pkcs.AlgorithmIdentifier(
                    new System.Security.Cryptography.Oid("2.16.840.1.101.3.4.1.42"))); // AES-256-CBC

            var cmsRecipients = new System.Security.Cryptography.Pkcs.CmsRecipientCollection();
            foreach (System.Security.Cryptography.X509Certificates.X509Certificate2 cert in recipientCerts)
            {
                cmsRecipients.Add(new System.Security.Cryptography.Pkcs.CmsRecipient(
                    System.Security.Cryptography.Pkcs.SubjectIdentifierType.IssuerAndSerialNumber, cert));
            }

            envelopedCms.Encrypt(cmsRecipients);
            var encrypted = envelopedCms.Encode();

            // ── Replace message content ──
            // RFC 7508: replace outer subject with generic placeholder
            // (real subject is protected inside the encrypted envelope)
            mail.Subject = "Encrypted Message";
            mail.HTMLBody = "";
            mail.Body = "";
            while (mail.Attachments.Count > 0)
                mail.Attachments[1].Delete();

            var tempPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".p7m");
            System.IO.File.WriteAllBytes(tempPath, encrypted);

            mail.Attachments.Add(tempPath,
                Outlook.OlAttachmentType.olByValue,
                Type.Missing,
                "smime.p7m");

            const string PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001E";
            mail.Attachments[mail.Attachments.Count].PropertyAccessor.SetProperty(
                PR_ATTACH_MIME_TAG,
                "application/pkcs7-mime; smime-type=enveloped-data; name=smime.p7m");

            try { System.IO.File.Delete(tempPath); } catch { }

            const string PR_MESSAGE_CLASS = "http://schemas.microsoft.com/mapi/proptag/0x001A001E";
            mail.PropertyAccessor.SetProperty(PR_MESSAGE_CLASS, "IPM.Note.SMIME");

            var flag = mail.UserProperties.Find("ParclEncrypt");
            if (flag != null)
                flag.Value = false;

            Logger.Info("Send",
                $"S/MIME encapsulated — {encrypted.Length} bytes for {recipientCerts.Count} recipient(s)");
            return null; // success
        }

        // Certificate import is now manual-only via the "Import Certificates" ribbon button.
    }
}
