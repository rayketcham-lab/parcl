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

            try
            {
                // ── Signing (deferred from compose toggle or auto-sign) ──
                bool shouldSign = false;

                // Check if user toggled Sign on this message
                var signFlag = mail.UserProperties.Find("ParclSign");
                if (signFlag != null && (bool)signFlag.Value)
                    shouldSign = true;

                // Check auto-sign setting
                if (Settings.Crypto.AlwaysSign &&
                    !string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                {
                    var signingCert = CertStore.FindByThumbprint(
                        Settings.UserProfile.SigningCertThumbprint!);
                    if (signingCert != null && signingCert.HasPrivateKey)
                        shouldSign = true;
                    else
                        Logger.Warn("Send",
                            "Signing certificate not found or has no private key — skipping");
                }

                if (shouldSign)
                {
                    const string PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
                    const int SECFLAG_SIGNED = 0x02;

                    var pa = mail.PropertyAccessor;
                    int flags;
                    try { flags = (int)pa.GetProperty(PR_SECURITY_FLAGS); }
                    catch { flags = 0; }

                    pa.SetProperty(PR_SECURITY_FLAGS, flags | SECFLAG_SIGNED);
                    Logger.Info("Send", "S/MIME signature flag applied at send time");
                }

                // ── Encryption (Parcl encapsulation at send time) ──
                bool shouldEncrypt = false;

                // Check if user toggled Encrypt on this message
                var encryptFlag = mail.UserProperties.Find("ParclEncrypt");
                if (encryptFlag != null && (bool)encryptFlag.Value)
                    shouldEncrypt = true;

                // Check auto-encrypt setting
                if (Settings.Crypto.AlwaysEncrypt)
                    shouldEncrypt = true;

                if (shouldEncrypt)
                {
                    Logger.Info("Send", "Encapsulating message in S/MIME envelope");
                    bool success = EncapsulateMessage(mail);
                    if (!success)
                    {
                        if (Settings.Behavior.PromptOnMissingCert)
                        {
                            var result = MessageBox.Show(
                                "Encryption certificates could not be found for all recipients. " +
                                "Send unencrypted?",
                                "Parcl — Missing Certificates",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Warning);
                            if (result == DialogResult.No)
                            {
                                cancel = true;
                                Logger.Info("Send",
                                    "User cancelled send due to missing recipient certificates");
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Send", "Failed during send processing", ex);
                MessageBox.Show(
                    $"Parcl encountered an error during send processing:\n{ex.Message}",
                    "Parcl Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Performs the actual S/MIME encapsulation at send time.
        /// Builds MIME from body + attachments, encrypts into CMS envelope,
        /// replaces message content with IPM.Note.SMIME + smime.p7m.
        /// Returns false if certs are missing for any recipient.
        /// </summary>
        private bool EncapsulateMessage(Outlook.MailItem mail)
        {
            // Resolve all recipient certs
            var recipientCerts =
                new System.Security.Cryptography.X509Certificates.X509Certificate2Collection();
            var missing = new System.Collections.Generic.List<string>();

            for (int i = 1; i <= mail.Recipients.Count; i++)
            {
                var recipient = mail.Recipients[i];
                var smtpAddr = ResolveSmtpAddress(recipient);
                var cert = ResolveRecipientCert(smtpAddr, recipient);
                if (cert != null)
                    recipientCerts.Add(cert);
                else
                    missing.Add(smtpAddr);
            }

            if (missing.Count > 0)
            {
                Logger.Warn("Send",
                    $"Missing certificates for: {string.Join(", ", missing)}");
                return false;
            }

            // Also encrypt to self so Sent Items are readable
            if (!string.IsNullOrEmpty(Settings.UserProfile.EncryptionCertThumbprint))
            {
                var selfCert = CertStore.FindByThumbprint(
                    Settings.UserProfile.EncryptionCertThumbprint!);
                if (selfCert != null)
                    recipientCerts.Add(selfCert);
            }

            // 1. Build MIME content from body + attachments
            var attachments = new System.Collections.Generic.List<Parcl.Core.Crypto.MimeAttachment>();
            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var att = mail.Attachments[i];
                var tempAtt = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + System.IO.Path.GetExtension(att.FileName));
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
                mail.Body, mail.HTMLBody, attachments.Count > 0 ? attachments : null);

            Logger.Debug("Send",
                $"MIME content built: {mimeContent.Length} bytes, {attachments.Count} attachment(s)");

            // 2. Encrypt MIME into CMS envelope
            var encrypted = SmimeHandler.Encrypt(mimeContent, recipientCerts);

            // 3. Clear original body and attachments
            mail.HTMLBody = "";
            mail.Body = "";
            while (mail.Attachments.Count > 0)
                mail.Attachments[1].Delete();

            // 4. Store encrypted CMS blob
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

            // 5. Set message class to S/MIME
            const string PR_MESSAGE_CLASS = "http://schemas.microsoft.com/mapi/proptag/0x001A001E";
            mail.PropertyAccessor.SetProperty(PR_MESSAGE_CLASS, "IPM.Note.SMIME");

            // Clean up the user property flag
            var flag = mail.UserProperties.Find("ParclEncrypt");
            if (flag != null)
                flag.Value = false;

            Logger.Info("Send",
                $"S/MIME encapsulated — {encrypted.Length} bytes for {recipientCerts.Count} recipient(s)");
            return true;
        }

        // Certificate import is now manual-only via the "Import Certificates" ribbon button.
    }
}
