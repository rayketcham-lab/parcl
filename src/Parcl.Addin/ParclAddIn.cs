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

                // Force Outlook's native signing to use SHA-256+ by patching the
                // MAPI profile binary. Parcl is authoritative over all S/MIME settings.
                // NOTE: Outlook loads S/MIME settings into memory BEFORE add-ins load.
                // If we patch the blob, we must restart Outlook for it to take effect.
                if (ForceOutlookSigningAlgorithm())
                {
                    // Blob was patched — Outlook needs to restart to pick up the change.
                    // Restart automatically: close Outlook and relaunch.
                    Logger.Info("AddIn", "Restarting Outlook to apply signing algorithm change...");
                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        await System.Threading.Tasks.Task.Delay(2000); // let add-in finish loading
                        try
                        {
                            var outlookPath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                            if (!string.IsNullOrEmpty(outlookPath))
                            {
                                System.Diagnostics.Process.Start(outlookPath);
                            }
                            _application?.Quit();
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn("AddIn", $"Auto-restart failed: {ex.Message}");
                        }
                    });
                }

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

                // Health indicator: warn user that encryption enforcement is offline
                try
                {
                    MessageBox.Show(
                        "Parcl failed to initialize.\n\n" +
                        $"Why: {ex.Message}\n\n" +
                        "Impact: Email encryption and signing are NOT active. " +
                        "Messages will be sent unencrypted until this is resolved.\n\n" +
                        "Fix: Restart Outlook. If this persists, reinstall Parcl or check the log at " +
                        "%APPDATA%\\Parcl\\logs\\",
                        "Parcl — Startup Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch { }
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
        private const string PR_ICON_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x10800003";
        // Icon indices matching what Outlook natively assigns to S/MIME messages
        // (verified from James/Dean's natively encrypted/signed messages)
        private const int ICON_ENCRYPTED = 275;        // padlock (Outlook native encrypted)
        private const int ICON_SIGNED = 276;            // ribbon/seal (Outlook native signed)
        private const int ICON_SIGNED_ENCRYPTED = 275;  // padlock (Outlook native enc+sign)

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
                    $"Classified: {(isEncrypted ? "encrypted" : "")}" +
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
                // ── Determine what's requested via SendDecision ──
                var parclEncryptFlag = false;
                var parclSignFlag = false;

                var encryptProp = mail.UserProperties.Find("ParclEncrypt");
                if (encryptProp != null && (bool)encryptProp.Value)
                    parclEncryptFlag = true;

                var signProp = mail.UserProperties.Find("ParclSign");
                if (signProp != null && (bool)signProp.Value)
                    parclSignFlag = true;

                bool hasSigningCert = false;
                if (!string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                {
                    var signingCert = CertStore.FindByThumbprint(
                        Settings.UserProfile.SigningCertThumbprint!);
                    hasSigningCert = signingCert != null && signingCert.HasPrivateKey;
                }

                var decision = Parcl.Core.Crypto.SendDecision.Evaluate(
                    parclEncryptFlag,
                    parclSignFlag,
                    Settings.Crypto.AlwaysEncrypt,
                    Settings.Crypto.AlwaysSign,
                    hasSigningCert,
                    Settings.Crypto.UseNativeSmime);

                bool shouldEncrypt = decision.ShouldEncrypt;
                bool shouldSign = decision.ShouldSign;

                if (decision.EncryptSource != null)
                    Logger.Debug("Send", $"Encrypt enabled by {decision.EncryptSource}");
                if (decision.SignSource != null)
                    Logger.Debug("Send", $"Sign enabled by {decision.SignSource}");

                shouldEncryptRequested = shouldEncrypt;
                Logger.Info("Send", $"Send mode: encrypt={shouldEncrypt}, sign={shouldSign}");

                // ── Apply ──
                // Parcl is authoritative for ALL signing — never delegate to Outlook's
                // PR_SECURITY_FLAGS for signing (it defaults to SHA-1).
                // If encrypting + signing: sign goes INSIDE the encrypted envelope (RFC 5751).
                // If only signing: Parcl signs and attaches as opaque SignedCms .p7m.

                if (shouldEncrypt)
                {
                    if (Settings.Crypto.UseNativeSmime)
                    {
                        // ── Native S/MIME via PR_SECURITY_FLAGS ──
                        // Check if all recipients have certs that Outlook can natively find.
                        // If any recipient has a cert mismatch (email != SMTP), Outlook will
                        // show the "Encryption Problems" dialog. In that case, fall back to
                        // Parcl envelope mode which handles cert mismatches gracefully.
                        Logger.Info("Send", "Checking native S/MIME compatibility"
                            + (shouldSign ? " (sign + encrypt)" : " (encrypt only)"));

                        bool allNativeCompatible = true;
                        for (int i = 1; i <= mail.Recipients.Count; i++)
                        {
                            var recipient = mail.Recipients[i];
                            var smtpAddr = ResolveSmtpAddress(recipient);
                            var cert = ResolveRecipientCert(smtpAddr, recipient);

                            if (cert == null || cert.NotAfter <= DateTime.UtcNow)
                            {
                                Logger.Warn("Send", $"No valid cert for {ParclLogger.SanitizeEmail(smtpAddr)}");
                                allNativeCompatible = false;
                                continue;
                            }

                            // Check if the cert email matches the SMTP address
                            // If it doesn't, Outlook's native engine won't find it
                            bool certMatchesSmtp = false;
                            var certEmail = Parcl.Core.Models.CertificateInfo.FromX509(cert).Email;
                            if (!string.IsNullOrEmpty(certEmail) &&
                                certEmail.Equals(smtpAddr, StringComparison.OrdinalIgnoreCase))
                            {
                                certMatchesSmtp = true;
                            }

                            // Also check if Subject contains the email
                            if (!certMatchesSmtp && cert.Subject != null &&
                                cert.Subject.ToLowerInvariant().Contains(smtpAddr.ToLowerInvariant()))
                            {
                                certMatchesSmtp = true;
                            }

                            if (certMatchesSmtp)
                            {
                                // Publish cert to AddressEntry for good measure
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
                                    }
                                }
                                catch { }

                                Logger.Info("Send",
                                    $"Native compatible: {ParclLogger.SanitizeEmail(smtpAddr)}");
                            }
                            else
                            {
                                Logger.Info("Send",
                                    $"Cert mismatch for {ParclLogger.SanitizeEmail(smtpAddr)} " +
                                    $"(cert={ParclLogger.SanitizeEmail(certEmail ?? "none")}) — will use Parcl envelope");
                                allNativeCompatible = false;
                            }
                        }

                        if (!allNativeCompatible)
                        {
                            // Fall back to Parcl envelope for cert-mismatched recipients
                            Logger.Info("Send", "Falling back to Parcl envelope (cert mismatch detected)");
                            string? encryptError = EncapsulateMessage(mail, shouldSign);
                            if (encryptError != null)
                            {
                                cancel = true;
                                Logger.Error("Send", $"Encryption failed — send blocked: {encryptError}");
                                MessageBox.Show(
                                    $"Message NOT sent — encryption failed.\n\n{encryptError}\n\n" +
                                    "To send anyway: toggle the Encrypt button off, then click Send again.",
                                    "Parcl — Send Blocked",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                                return;
                            }
                        }
                        else
                        {
                            // All recipients native-compatible — use PR_SECURITY_FLAGS
                            // Signing algorithm is controlled via registry (ConfigureOutlookSigningAlgorithm)
                            Logger.Info("Send", "All recipients native-compatible — using PR_SECURITY_FLAGS");

                            const string PR_SEC = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
                            var pa = mail.PropertyAccessor;
                            int flags;
                            try { flags = (int)pa.GetProperty(PR_SEC); }
                            catch { flags = 0; }

                            flags |= 0x01; // SECFLAG_ENCRYPTED
                            if (shouldSign)
                                flags |= 0x02; // SECFLAG_SIGNED (uses SHA-256+ via registry)

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
                                $"Message NOT sent — encryption failed.\n\n{encryptError}\n\n" +
                                "To send anyway: toggle the Encrypt button off, then click Send again.",
                                "Parcl — Send Blocked",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                else if (shouldSign)
                {
                    // Sign-only: use Outlook's native signing (PR_SECURITY_FLAGS) for proper
                    // inline display in the reading pane.
                    //
                    // IMPORTANT: The signing hash algorithm is controlled by Outlook's Trust Center
                    // Email Security settings (not registry, not per-message). Parcl cannot change
                    // this programmatically. If the user has SHA-1 configured, we detect it by
                    // checking the last sent signed message and warn them to change it.
                    const string PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
                    const int SECFLAG_SIGNED = 0x02;

                    var pa = mail.PropertyAccessor;
                    int flags;
                    try { flags = (int)pa.GetProperty(PR_SECURITY_FLAGS); }
                    catch { flags = 0; }

                    pa.SetProperty(PR_SECURITY_FLAGS, flags | SECFLAG_SIGNED);
                    Logger.Info("Send", "Sign-only via native S/MIME");
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
                        $"Message NOT sent — an unexpected error occurred during encryption.\n\n" +
                        $"Why: {ex.Message}\n\n" +
                        "Fix: Check that your certificates are valid in Parcl > Select Certificates.\n" +
                        "To send without encryption: toggle the Encrypt button off, then click Send again.",
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
                    errors.Add($"{smtpAddr}: No encryption certificate found.\n" +
                        "  Why: Parcl could not locate an S/MIME certificate for this recipient.\n" +
                        "  Fix: Open Parcl > Contacts and import their certificate, or ask them to send it via Certificate Exchange.");
                    Logger.Warn("Send", $"No certificate found for {ParclLogger.SanitizeEmail(smtpAddr)}");
                    continue;
                }

                // Check expiry
                if (cert.NotAfter <= DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate expired on {cert.NotAfter:yyyy-MM-dd}.\n" +
                        "  Why: The recipient's S/MIME certificate is past its validity period.\n" +
                        "  Fix: Ask the recipient to renew their certificate and send you the updated one via Certificate Exchange.");
                    Logger.Warn("Send", $"Certificate expired for {ParclLogger.SanitizeEmail(smtpAddr)} (expired {cert.NotAfter:yyyy-MM-dd})");
                    continue;
                }

                if (cert.NotBefore > DateTime.UtcNow)
                {
                    errors.Add($"{smtpAddr}: Certificate not yet valid (starts {cert.NotBefore:yyyy-MM-dd}).\n" +
                        "  Why: The recipient's certificate has a future start date and cannot be used yet.\n" +
                        "  Fix: Wait until {cert.NotBefore:yyyy-MM-dd}, or ask the recipient for a currently valid certificate.");
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
                return "No recipients have valid certificates.\n\n" +
                    "Why: Parcl needs an S/MIME certificate for each recipient to encrypt the message.\n\n" +
                    "Fix: Open Parcl > Contacts to import certificates, or use Certificate Exchange to request them.";

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
                    contentToEncrypt = SmimeHandler.Sign(mimeContent, signingCert);
                    Logger.Info("Send",
                        $"Signed: {mimeContent.Length} bytes MIME -> {contentToEncrypt.Length} bytes SignedCms ({Settings.Crypto.HashAlgorithm})");
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

        /// <summary>
        /// Forces Outlook's S/MIME security profile to use the configured hash algorithm.
        /// Directly patches the binary blob in the Outlook MAPI profile registry.
        /// The hash algorithm index is stored as a DWORD at offset 0x28 in the
        /// S/MIME settings property (11020355) under the Outlook profile key
        /// c02ebc5353d9cd11975200aa004ae40e.
        ///
        /// Algorithm indices (1-based in the DER algorithm list):
        ///   7  = SHA-1    (BLOCKED — insecure)
        ///   8  = SHA-512
        ///   9  = SHA-384
        ///   10 = SHA-256
        /// </summary>
        // DER-encoded hash algorithm OIDs for S/MIME profile reordering
        private static readonly byte[] DerSHA1   = { 0x30, 0x07, 0x06, 0x05, 0x2B, 0x0E, 0x03, 0x02, 0x1A };
        private static readonly byte[] DerSHA256 = { 0x30, 0x0B, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01 };
        private static readonly byte[] DerSHA384 = { 0x30, 0x0B, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x02 };
        private static readonly byte[] DerSHA512 = { 0x30, 0x0B, 0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x03 };

        /// <summary>
        /// Forces Outlook's S/MIME security profile to use the configured hash algorithm.
        /// Outlook picks the FIRST hash algorithm in the DER capability list stored in the
        /// MAPI profile binary blob. This method reorders the hash OIDs so the configured
        /// algorithm (SHA-256 by default) is first.
        ///
        /// The S/MIME settings are in registry property 11020355 under the Outlook profile
        /// key c02ebc5353d9cd11975200aa004ae40e. The DER SEQUENCE at the end of the blob
        /// contains encryption algorithms followed by hash algorithms.
        /// </summary>
        /// <returns>true if the blob was patched (Outlook restart needed), false if already correct.</returns>
        private bool ForceOutlookSigningAlgorithm()
        {
            bool patched = false;
            try
            {
                var targetAlgo = Settings.Crypto.HashAlgorithm ?? "SHA-256";

                // Determine the desired hash order: target algorithm first, then others, SHA-1 last
                byte[] targetDer;
                byte[][] otherDers;
                switch (targetAlgo.ToUpperInvariant())
                {
                    case "SHA-384":
                        targetDer = DerSHA384;
                        otherDers = new[] { DerSHA256, DerSHA512, DerSHA1 };
                        break;
                    case "SHA-512":
                        targetDer = DerSHA512;
                        otherDers = new[] { DerSHA256, DerSHA384, DerSHA1 };
                        break;
                    default: // SHA-256
                        targetDer = DerSHA256;
                        otherDers = new[] { DerSHA512, DerSHA384, DerSHA1 };
                        break;
                }

                const string profileBase = @"Software\Microsoft\Office\16.0\Outlook\Profiles";
                using (var profilesKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(profileBase))
                {
                    if (profilesKey == null) return false;

                    foreach (var profileName in profilesKey.GetSubKeyNames())
                    {
                        var smimeKeyPath = $@"{profileBase}\{profileName}\c02ebc5353d9cd11975200aa004ae40e";
                        using (var smimeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(smimeKeyPath, writable: true))
                        {
                            if (smimeKey == null) continue;

                            var blob = smimeKey.GetValue("11020355") as byte[];
                            if (blob == null || blob.Length < 0x170) continue;

                            // Find the hash algorithm section in the blob.
                            // SHA-1 DER starts with 30 07 06 05 2B 0E 03 02 1A.
                            // SHA-256/384/512 DER starts with 30 0B 06 09 60 86 48 01 65 03 04 02 XX.
                            // The hash section is 48 bytes (9+13+13+13) and follows the encryption OIDs.
                            int hashStart = FindSequence(blob, DerSHA1);
                            if (hashStart < 0)
                            {
                                // SHA-1 might already be moved — look for SHA-256 as start
                                hashStart = FindSequence(blob, DerSHA256);
                                if (hashStart < 0)
                                {
                                    Logger.Debug("AddIn", "Could not locate hash algorithms in S/MIME profile");
                                    continue;
                                }
                            }

                            // Check if the target is already first
                            bool alreadyFirst = true;
                            for (int j = 0; j < targetDer.Length && (hashStart + j) < blob.Length; j++)
                            {
                                if (blob[hashStart + j] != targetDer[j]) { alreadyFirst = false; break; }
                            }

                            if (alreadyFirst)
                            {
                                Logger.Debug("AddIn",
                                    $"Outlook signing algorithm already {targetAlgo} (first in DER list)");
                                continue;
                            }

                            // Build new hash section: target first, then others
                            var newHashSection = new byte[48]; // 13+13+13+9 = 48
                            int pos = 0;
                            Array.Copy(targetDer, 0, newHashSection, pos, targetDer.Length);
                            pos += targetDer.Length;
                            foreach (var other in otherDers)
                            {
                                Array.Copy(other, 0, newHashSection, pos, other.Length);
                                pos += other.Length;
                            }

                            // Find the start of the hash section (first hash OID in the blob)
                            // The 4 hash OIDs are contiguous and total 48 bytes
                            int sectionStart = Math.Min(
                                hashStart,
                                Math.Min(
                                    FindSequence(blob, DerSHA256) >= 0 ? FindSequence(blob, DerSHA256) : int.MaxValue,
                                    Math.Min(
                                        FindSequence(blob, DerSHA384) >= 0 ? FindSequence(blob, DerSHA384) : int.MaxValue,
                                        FindSequence(blob, DerSHA512) >= 0 ? FindSequence(blob, DerSHA512) : int.MaxValue
                                    )
                                )
                            );

                            // Write new hash section
                            Array.Copy(newHashSection, 0, blob, sectionStart, newHashSection.Length);

                            smimeKey.SetValue("11020355", blob, Microsoft.Win32.RegistryValueKind.Binary);
                            patched = true;

                            Logger.Info("AddIn",
                                $"Outlook signing algorithm forced to {targetAlgo} (reordered DER hash list, first OID is now {targetAlgo})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("AddIn", $"Could not force Outlook signing algorithm: {ex.Message}");
            }
            return patched;
        }

        private static int FindSequence(byte[] haystack, byte[] needle)
        {
            for (int i = 0; i <= haystack.Length - needle.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < needle.Length; j++)
                {
                    if (haystack[i + j] != needle[j]) { match = false; break; }
                }
                if (match) return i;
            }
            return -1;
        }

        /// <summary>
        /// Signs a message using Parcl's own SmimeHandler (not Outlook native).
        /// Builds MIME content from the mail, signs with the configured hash algorithm,
        /// and replaces the message with an opaque signed .p7m attachment.
        /// Returns null on success, or an error message if signing failed.
        /// </summary>
        private string? SignOnlyMessage(Outlook.MailItem mail)
        {
            if (string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                return "No signing certificate configured.\n\n" +
                    "Fix: Go to Parcl > Select Certificates and choose a signing certificate.";

            var signingCert = CertStore.FindByThumbprint(Settings.UserProfile.SigningCertThumbprint!);
            if (signingCert == null || !signingCert.HasPrivateKey)
                return "Signing certificate not found or missing private key.\n\n" +
                    "Fix: Go to Parcl > Select Certificates and verify your signing certificate is installed with a private key.";

            Logger.Debug("Send",
                $"Sign-only cert: {signingCert.Subject}, thumbprint: {signingCert.Thumbprint.Substring(0, 8)}");

            // ── Build MIME content from the mail ──
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

            var mimeContent = Parcl.Core.Crypto.MimeBuilder.Build(
                mail.Body, mail.HTMLBody,
                attachments.Count > 0 ? attachments : null);

            Logger.Debug("Send",
                $"Sign-only MIME built: {mimeContent.Length} bytes, {attachments.Count} attachment(s)");

            // ── Sign with SmimeHandler (uses configured hash algorithm) ──
            byte[] signedData;
            try
            {
                signedData = SmimeHandler.Sign(mimeContent, signingCert);
            }
            catch (Exception ex)
            {
                Logger.Error("Send", "SmimeHandler.Sign() failed", ex);
                return $"Signing failed: {ex.Message}";
            }

            Logger.Info("Send",
                $"Signed: {mimeContent.Length} bytes MIME -> {signedData.Length} bytes SignedCms ({Settings.Crypto.HashAlgorithm})");

            // ── Replace message content with signed .p7m ──
            while (mail.Attachments.Count > 0)
                mail.Attachments[1].Delete();

            mail.HTMLBody = "<div style=\"font-family:Segoe UI,sans-serif;padding:24px;\">" +
                "<h3 style=\"color:#4FC3F7;\">&#9997; Signed with Parcl</h3>" +
                "<p>This message is digitally signed. Use the <b>Parcl Decrypt</b> button on the ribbon to verify and read it.</p>" +
                "<p style=\"color:#888;font-size:11px;\">If you don't have Parcl installed, " +
                "the smime.p7m attachment can be verified by any S/MIME-compatible email client.</p>" +
                "</div>";

            var tempPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".p7m");
            System.IO.File.WriteAllBytes(tempPath, signedData);

            mail.Attachments.Add(tempPath,
                Outlook.OlAttachmentType.olByValue,
                Type.Missing,
                "smime.p7m");

            try { System.IO.File.Delete(tempPath); } catch { }

            var flag = mail.UserProperties.Find("ParclSign");
            if (flag != null)
                flag.Value = false;

            Logger.Info("Send",
                $"Sign-only complete — {signedData.Length} bytes signed with {Settings.Crypto.HashAlgorithm}");
            return null; // success
        }

        // Certificate import is now manual-only via the "Import Certificates" ribbon button.
    }
}
