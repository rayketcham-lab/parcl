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

            // FIPS 140-2 compliance check
            try
            {
                var fipsEnabled = System.Security.Cryptography.CryptoConfig.AllowOnlyFipsAlgorithms;
                Logger.Info("AddIn", $"FIPS mode: {(fipsEnabled ? "ENABLED" : "disabled")}");
            }
            catch { }

            try
            {
                Settings = ParclSettings.Load();

                // Apply log level from settings
                if (Enum.TryParse<Parcl.Core.Config.LogLevel>(
                        Settings.Behavior.LogLevel, true, out var configuredLevel))
                    Logger.SetMinLevel(configuredLevel);

                Logger.Debug("AddIn",
                    $"Settings loaded — {Settings.LdapDirectories.Count} LDAP directories configured");

                CertStore = new CertificateStore();
                SmimeHandler = new SmimeHandler(
                    Settings.Crypto.EncryptionAlgorithm,
                    Settings.Crypto.HashAlgorithm);
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
                {
                    explorer.SelectionChange += Explorer_SelectionChange;
                    explorer.FolderSwitch += Explorer_FolderSwitch;
                }

                // Also catch new inspectors (opened messages)
                _application.Inspectors.NewInspector += Inspectors_NewInspector;

                // Hook new mail for inbox icon classification
                _application.NewMailEx += Application_NewMailEx;

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
            catch (Exception ex) { Logger?.Debug("Ribbon", $"Invalidate failed on SelectionChange: {ex.Message}"); }

            // Classify selected message icon if not already set
            try
            {
                var explorer = _application?.ActiveExplorer();
                if (explorer?.Selection?.Count > 0 && explorer.Selection[1] is Outlook.MailItem mail)
                {
                    int currentIcon = -1;
                    try { currentIcon = (int)mail.PropertyAccessor.GetProperty(PR_ICON_INDEX); }
                    catch { }

                    // Only classify if icon not yet set (icon == -1 or default 256/272)
                    if (currentIcon == -1 || currentIcon == 256 || currentIcon == 272)
                        ClassifyAndSetIcon(mail);
                }
            }
            catch { }
        }

        private void Explorer_FolderSwitch()
        {
            // Hide/show Parcl tab when switching between Mail/Calendar/People
            try { _ribbon?.Invalidate(); }
            catch (Exception ex) { Logger?.Debug("Ribbon", $"Invalidate failed on FolderSwitch: {ex.Message}"); }
        }

        // ── Inbox icon classification ─────────────────────────────────────
        // PR_ICON_INDEX values: Outlook built-in S/MIME icons
        private const string PR_ICON_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x10800003";
        private const int ICON_ENCRYPTED = 1604;     // padlock
        private const int ICON_SIGNED = 1603;         // ribbon/seal
        private const int ICON_SIGNED_ENCRYPTED = 1605; // padlock + ribbon

        private void Application_NewMailEx(string entryIDCollection)
        {
            try
            {
                foreach (var entryID in entryIDCollection.Split(','))
                {
                    var id = entryID.Trim();
                    if (string.IsNullOrEmpty(id)) continue;

                    var item = _application!.Session.GetItemFromID(id);
                    if (item is Outlook.MailItem mail)
                    {
                        ClassifyAndSetIcon(mail);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger?.Debug("Icons", $"NewMailEx icon classification failed: {ex.Message}");
            }
        }

        private void ClassifyAndSetIcon(Outlook.MailItem mail)
        {
            try
            {
                bool isEncrypted = false;
                bool isSigned = false;

                // Check for .p7m attachment (Parcl-encrypted)
                for (int i = 1; i <= mail.Attachments.Count; i++)
                {
                    if (mail.Attachments[i].FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                    {
                        isEncrypted = true;
                        break;
                    }
                }

                // Check PR_SECURITY_FLAGS (native S/MIME)
                try
                {
                    var flags = (int)mail.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x6E010003");
                    if ((flags & 0x01) != 0) isEncrypted = true;
                    if ((flags & 0x02) != 0) isSigned = true;
                }
                catch { }

                // Check message class
                try
                {
                    if (mail.MessageClass == "IPM.Note.SMIME")
                        isEncrypted = true;
                    if (mail.MessageClass == "IPM.Note.SMIME.MultipartSigned")
                        isSigned = true;
                }
                catch { }

                if (!isEncrypted && !isSigned) return;

                int iconIndex = isEncrypted && isSigned ? ICON_SIGNED_ENCRYPTED
                              : isEncrypted ? ICON_ENCRYPTED
                              : ICON_SIGNED;

                mail.PropertyAccessor.SetProperty(PR_ICON_INDEX, iconIndex);
                mail.Save();

                Logger?.Debug("Icons",
                    $"Icon set: {(isEncrypted ? "encrypted" : "")}" +
                    $"{(isSigned ? " signed" : "")} -> icon {iconIndex} for: {mail.Subject}");
            }
            catch (Exception ex) { Logger?.Debug("Icons", $"ClassifyAndSetIcon failed: {ex.Message}"); }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            // When a message is opened in its own window, refresh ribbon state
            try
            {
                inspector.Activate();
                _ribbon?.Invalidate();

                // Auto-decrypt if enabled
                if (Settings.Behavior.AutoDecrypt)
                {
                    if (inspector.CurrentItem is Outlook.MailItem mail && mail.Sent)
                    {
                        TryAutoDecrypt(mail);
                    }
                }
            }
            catch (Exception ex) { Logger?.Debug("Inspector", $"NewInspector handler failed: {ex.Message}"); }
        }

        /// <summary>
        /// Attempts to auto-decrypt a message if it has a .p7m attachment.
        /// Silently skips if no encrypted content or if decryption fails.
        /// </summary>
        private void TryAutoDecrypt(Outlook.MailItem mail)
        {
            try
            {
                // Find .p7m attachment
                Outlook.Attachment? p7m = null;
                for (int i = 1; i <= mail.Attachments.Count; i++)
                {
                    if (mail.Attachments[i].FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                    {
                        p7m = mail.Attachments[i];
                        break;
                    }
                }

                if (p7m == null) return;

                Logger.Info("AutoDecrypt", $"Auto-decrypting message: {mail.Subject ?? "(no subject)"}");

                var tempPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    System.IO.Path.GetRandomFileName() + ".p7m");
                p7m.SaveAsFile(tempPath);
                byte[] encryptedData;
                try { encryptedData = System.IO.File.ReadAllBytes(tempPath); }
                finally { try { System.IO.File.Delete(tempPath); } catch { } }

                var result = SmimeHandler.Decrypt(encryptedData);
                if (!result.Success || result.Content == null)
                {
                    Logger.Debug("AutoDecrypt", $"Could not auto-decrypt: {result.ErrorMessage}");
                    return;
                }

                byte[] mimeBytes = result.Content;
                Logger.Debug("AutoDecrypt", $"CMS decrypted: {mimeBytes.Length} bytes");

                // Unwrap SignedCms if present
                try
                {
                    var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
                    signedCms.Decode(mimeBytes);
                    signedCms.CheckSignature(verifySignatureOnly: false);
                    mimeBytes = signedCms.ContentInfo.Content;
                    Logger.Debug("AutoDecrypt", "Signature verified");
                }
                catch (System.Security.Cryptography.CryptographicException) { }

                // Parse MIME and restore body
                string mimeText = System.Text.Encoding.UTF8.GetString(mimeBytes);
                var headers = Parcl.Core.Crypto.MimeBuilder.ExtractProtectedHeaders(mimeText);
                var extracted = Parcl.Core.Crypto.MimeBuilder.ExtractBody(mimeText);

                if (extracted.HasContent)
                {
                    if (!string.IsNullOrEmpty(extracted.HtmlBody))
                        mail.HTMLBody = extracted.HtmlBody;
                    else if (!string.IsNullOrEmpty(extracted.TextBody))
                        mail.Body = extracted.TextBody;
                }

                if (headers?.Subject != null)
                    mail.Subject = headers.Subject;

                // Remove .p7m attachment
                for (int i = mail.Attachments.Count; i >= 1; i--)
                {
                    if (mail.Attachments[i].FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                        mail.Attachments[i].Delete();
                }

                // Restore envelope attachments
                foreach (var att in extracted.Attachments)
                {
                    var attTemp = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        System.IO.Path.GetRandomFileName() + "_" + att.FileName);
                    try
                    {
                        System.IO.File.WriteAllBytes(attTemp, att.Data);
                        mail.Attachments.Add(attTemp,
                            Outlook.OlAttachmentType.olByValue, Type.Missing, att.FileName);
                    }
                    finally { try { System.IO.File.Delete(attTemp); } catch { } }
                }

                mail.Save();
                Logger.Info("AutoDecrypt", $"Message auto-decrypted: {mail.Subject}");
            }
            catch (Exception ex)
            {
                Logger.Debug("AutoDecrypt", $"Auto-decrypt skipped: {ex.Message}");
            }
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
            Logger.Info("Send", $"ItemSend intercepted — to: {ParclLogger.SanitizeEmail(mail.To)}, subject: {subjectPreview}");

            bool shouldEncryptRequested = false;
            try
            {
                // ── Determine what's requested ──
                bool shouldSign = false;
                bool shouldEncrypt = false;

                var signFlag = mail.UserProperties.Find("ParclSign");
                if (signFlag != null && (bool)signFlag.Value)
                {
                    shouldSign = true;
                    Logger.Debug("Send", "Sign flag set by user toggle");
                }

                if (Settings.Crypto.AlwaysSign &&
                    !string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                {
                    var signingCert = CertStore.FindByThumbprint(
                        Settings.UserProfile.SigningCertThumbprint!);
                    if (signingCert != null && signingCert.HasPrivateKey)
                    {
                        shouldSign = true;
                        Logger.Debug("Send", "Sign enabled by AlwaysSign setting");
                    }
                }

                var encryptFlag = mail.UserProperties.Find("ParclEncrypt");
                if (encryptFlag != null && (bool)encryptFlag.Value)
                {
                    shouldEncrypt = true;
                    Logger.Debug("Send", "Encrypt flag set by user toggle");
                }

                if (Settings.Crypto.AlwaysEncrypt)
                {
                    shouldEncrypt = true;
                    Logger.Debug("Send", "Encrypt enabled by AlwaysEncrypt setting");
                }

                shouldEncryptRequested = shouldEncrypt;
                Logger.Info("Send", $"Send mode: encrypt={shouldEncrypt}, sign={shouldSign}");

                // ── Apply ──
                // If encrypting: do it ourselves (sign goes INSIDE the encrypted envelope per RFC 5751).
                // If only signing: use Outlook's native PR_SECURITY_FLAGS.
                // Never set PR_SECURITY_FLAGS for sign when also encrypting — that double-wraps
                // and the recipient sees "Signed" as the outer layer instead of "Encrypted".

                if (shouldEncrypt)
                {
                    if (Settings.Crypto.UseNativeSmime)
                    {
                        // ── Native S/MIME via PR_SECURITY_FLAGS ──
                        // Outlook handles CMS encryption internally and formats the MIME
                        // correctly for reading pane inline display. Publish certs to
                        // AddressEntry so Outlook can find them for recipients with
                        // RDN/email mismatches.
                        Logger.Info("Send", "Using native Outlook S/MIME"
                            + (shouldSign ? " (sign + encrypt)" : " (encrypt only)"));

                        // Publish certs to recipients so Outlook finds them
                        for (int i = 1; i <= mail.Recipients.Count; i++)
                        {
                            var recipient = mail.Recipients[i];
                            var smtpAddr = ResolveSmtpAddress(recipient);
                            var cert = ResolveRecipientCert(smtpAddr, recipient);
                            if (cert != null && cert.NotAfter > DateTime.UtcNow)
                            {
                                try
                                {
                                    var addrEntry = recipient.AddressEntry;
                                    if (addrEntry != null)
                                    {
                                        var certBytes = cert.Export(
                                            System.Security.Cryptography.X509Certificates.X509ContentType.Cert);
                                        addrEntry.PropertyAccessor.SetProperty(
                                            PR_USER_X509_CERT,
                                            new object[] { certBytes });
                                        Logger.Info("Send",
                                            $"Cert published for {ParclLogger.SanitizeEmail(smtpAddr)}: {cert.Subject}");
                                    }
                                }
                                catch (Exception pubEx)
                                {
                                    Logger.Debug("Send", $"Cert publish failed for {smtpAddr}: {pubEx.Message}");
                                }
                            }
                            else
                            {
                                Logger.Warn("Send", $"No valid cert for {ParclLogger.SanitizeEmail(smtpAddr)}");
                            }
                        }

                        // Set PR_SECURITY_FLAGS — Outlook encrypts at send time
                        const string PR_SEC = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
                        var pa = mail.PropertyAccessor;
                        int flags;
                        try { flags = (int)pa.GetProperty(PR_SEC); }
                        catch { flags = 0; }

                        flags |= 0x01; // SECFLAG_ENCRYPTED
                        if (shouldSign)
                            flags |= 0x02; // SECFLAG_SIGNED

                        pa.SetProperty(PR_SEC, flags);

                        // Clear Parcl flags — Outlook handles from here
                        var encFlag = mail.UserProperties.Find("ParclEncrypt");
                        if (encFlag != null) encFlag.Value = false;
                        var sigFlag = mail.UserProperties.Find("ParclSign");
                        if (sigFlag != null) sigFlag.Value = false;

                        Logger.Info("Send",
                            $"Native S/MIME: flags=0x{flags:X}, " +
                            $"encrypt=true, sign={shouldSign}");
                    }
                    else
                    {
                        // ── Parcl envelope: our own CMS encryption with protected headers ──
                        Logger.Info("Send", "Using Parcl S/MIME envelope"
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
        /// <summary>
        /// Builds CMS envelope and returns encrypted bytes without modifying the mail item.
        /// Returns null on error (error string in out param). Used by Extended MAPI path.
        /// </summary>
        private byte[]? BuildCmsEnvelope(Outlook.MailItem mail, bool alsoSign, out string? error)
        {
            error = null;
            // Reuse EncapsulateMessage logic but capture the encrypted bytes
            // before they get written to the message
            _lastCmsBytes = null;
            _captureCmsBytes = true;
            try
            {
                error = EncapsulateMessage(mail, alsoSign);
                return _lastCmsBytes;
            }
            finally
            {
                _captureCmsBytes = false;
                _lastCmsBytes = null;
            }
        }

        // Used by BuildCmsEnvelope to capture encrypted bytes mid-EncapsulateMessage
        private bool _captureCmsBytes;
        private byte[]? _lastCmsBytes;

        private string? EncapsulateMessage(Outlook.MailItem mail, bool alsoSign = false)
        {
            // ── Validate ALL recipients have valid certs ──
            Logger.Info("Send", $"Validating certificates for {mail.Recipients.Count} recipient(s)");
            var recipientCerts =
                new System.Security.Cryptography.X509Certificates.X509Certificate2Collection();
            var errors = new System.Collections.Generic.List<string>();

            for (int i = 1; i <= mail.Recipients.Count; i++)
            {
                var recipient = mail.Recipients[i];
                var smtpAddr = ResolveSmtpAddress(recipient);
                Logger.Debug("Send", $"Resolving cert for recipient {i}: {smtpAddr}");
                var cert = ResolveRecipientCert(smtpAddr, recipient);

                if (cert == null)
                {
                    errors.Add($"{smtpAddr}: No certificate found");
                    Logger.Warn("Send", $"No certificate found for {ParclLogger.SanitizeEmail(smtpAddr)}");
                    continue;
                }

                // Check expiry
                if (cert.NotAfter <= DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate expired on {cert.NotAfter:yyyy-MM-dd}");
                    Logger.Warn("Send", $"Certificate expired for {ParclLogger.SanitizeEmail(smtpAddr)} (expired {cert.NotAfter:yyyy-MM-dd})");
                    continue;
                }

                if (cert.NotBefore > DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate not yet valid (starts {cert.NotBefore:yyyy-MM-dd})");
                    Logger.Warn("Send", $"Certificate not yet valid for {ParclLogger.SanitizeEmail(smtpAddr)}");
                    continue;
                }

                Logger.Info("Send", $"Cert OK for {ParclLogger.SanitizeEmail(smtpAddr)}: {cert.Subject}, expires {cert.NotAfter:yyyy-MM-dd}");
                recipientCerts.Add(cert);
            }

            if (errors.Count > 0)
            {
                Logger.Error("Send",
                    $"Encryption blocked: {errors.Count} recipient(s) failed cert validation");
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
                {
                    recipientCerts.Add(selfCert);
                    Logger.Debug("Send", $"Added self-encrypt cert: {selfCert.Subject}");
                }
            }
            Logger.Info("Send", $"Encrypting to {recipientCerts.Count} certificate(s)");

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
                Logger.Info("Send", $"Signing MIME content inside envelope (sign-then-encrypt per RFC 5751)");
                var signingCert = CertStore.FindByThumbprint(
                    Settings.UserProfile.SigningCertThumbprint!);
                if (signingCert != null && signingCert.HasPrivateKey)
                {
                    Logger.Debug("Send", $"Signing cert: {signingCert.Subject}, thumbprint: {signingCert.Thumbprint.Substring(0, 8)}");
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
                        $"Signed: {mimeContent.Length} bytes MIME -> {contentToEncrypt.Length} bytes SignedCms (SHA-256)");
                }
                else
                {
                    Logger.Warn("Send", "Signing cert not found or missing private key, encrypting without signature");
                }
            }

            // ── Encrypt with AES-256-CBC ──
            Logger.Info("Send", $"Encrypting {contentToEncrypt.Length} bytes with AES-256-CBC");
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

            // Capture CMS bytes if BuildCmsEnvelope is calling us
            if (_captureCmsBytes)
                _lastCmsBytes = encrypted;

            // ── Replace message content ──
            // Replace outer subject with generic placeholder
            // (real subject is protected inside the encrypted envelope via RFC 7508)
            mail.Subject = "Encrypted Message";
            while (mail.Attachments.Count > 0)
                mail.Attachments[1].Delete();

            // Set a visible placeholder body so the recipient knows to use Parcl
            mail.HTMLBody = "<div style=\"font-family:Segoe UI,sans-serif;padding:24px;\">" +
                "<h3 style=\"color:#4FC3F7;\">&#128274; Encrypted with Parcl</h3>" +
                "<p>This message is encrypted. Use the <b>Parcl Decrypt</b> button on the ribbon to read it.</p>" +
                "<p style=\"color:#888;font-size:11px;\">If you don't have Parcl installed, ask the sender for the unencrypted version.</p>" +
                "</div>";

            var tempPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".p7m");
            System.IO.File.WriteAllBytes(tempPath, encrypted);

            mail.Attachments.Add(tempPath,
                Outlook.OlAttachmentType.olByValue,
                Type.Missing,
                "smime.p7m");

            try { System.IO.File.Delete(tempPath); } catch { }

            // DO NOT set IPM.Note.SMIME — Outlook/Exchange intercepts that class
            // during transport and strips the attachment, leaving an empty message.
            // Keep as IPM.Note so the smime.p7m attachment arrives intact.
            // Parcl handles decryption explicitly via the Decrypt ribbon button.

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
