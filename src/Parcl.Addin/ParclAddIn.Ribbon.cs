using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Parcl.Addin.Dialogs;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parcl.Addin
{
    public partial class ParclAddIn
    {
        private IRibbonUI? _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            Logger?.Debug("Ribbon", $"GetCustomUI called for ribbonID={ribbonID}");
            return GetResourceText("Parcl.Addin.ParclRibbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Logger?.Info("Ribbon", "Parcl ribbon loaded into Outlook");
        }

        public Bitmap GetButtonImage(string imageId)
        {
            Logger?.Debug("Ribbon", $"Loading custom icon for: {imageId}");
            return RibbonIcons.GetIcon(imageId, 32);
        }

        // ── MAPI property constants ──────────────────────────────────────────
        private const string PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003";
        private const string PR_USER_X509_CERT = "http://schemas.microsoft.com/mapi/proptag/0x3A701102";
        private const int SECFLAG_ENCRYPTED = 0x01;
        private const int SECFLAG_SIGNED = 0x02;

        // ── Encrypt toggle (ribbon) ──────────────────────────────────────────
        public void OnEncryptToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null) return;

                if (!mail.Sent)
                {
                    if (pressed)
                        EncryptOutgoing(mail);
                    else
                        RemoveFlag(mail, SECFLAG_ENCRYPTED, "Encrypt");
                }
                else if (pressed)
                {
                    EncryptAtRest(mail);
                }

                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Encrypt", "Encryption toggle failed", ex);
                MessageBox.Show($"Encryption error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool GetEncryptPressed(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItemSafe(control);
                if (mail == null) return false;

                // Check our user property flag (compose mode)
                var flag = mail.UserProperties.Find("ParclEncrypt");
                if (flag != null && (bool)flag.Value)
                    return true;

                // Check if already encapsulated (sent/received S/MIME)
                try
                {
                    if (mail.MessageClass == "IPM.Note.SMIME")
                        return true;
                }
                catch { }

                return false;
            }
            catch { return false; }
        }

        // ── Encrypt button (context menu — same logic, no bool) ──────────────
        public void OnEncryptClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null) return;

                if (!mail.Sent)
                    EncryptOutgoing(mail);
                else
                    EncryptAtRest(mail);

                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Encrypt", "Encryption action failed", ex);
                MessageBox.Show($"Encryption error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Sign toggle (ribbon) ────────────────────────────────────────────
        public void OnSignToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || mail.Sent) return;

                if (pressed)
                    SignOutgoing(mail);
                else
                    RemoveFlag(mail, SECFLAG_SIGNED, "Sign");

                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Sign", "Signing toggle failed", ex);
                MessageBox.Show($"Signing error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool GetSignPressed(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItemSafe(control);
                if (mail == null) return false;

                // Check our user property flag (compose mode)
                var flag = mail.UserProperties.Find("ParclSign");
                if (flag != null && (bool)flag.Value)
                    return true;

                // Check Outlook-native S/MIME signed flag (received/sent messages)
                if ((GetSecurityFlags(mail.PropertyAccessor) & SECFLAG_SIGNED) != 0)
                    return true;

                // Check message class for signed messages
                try
                {
                    if (mail.MessageClass == "IPM.Note.SMIME.MultipartSigned" ||
                        mail.MessageClass == "IPM.Note.SMIME")
                        return true;
                }
                catch { }

                return false;
            }
            catch { return false; }
        }

        // ── Sign button (context menu) ──────────────────────────────────────
        public void OnSignClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || mail.Sent) return;
                SignOutgoing(mail);
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Sign", "Signing action failed", ex);
                MessageBox.Show($"Signing error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Remove flags (context menu) ─────────────────────────────────────
        public void OnRemoveEncryptionClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || mail.Sent) return;
                RemoveFlag(mail, SECFLAG_ENCRYPTED, "Encrypt");
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Encrypt", "Remove encryption failed", ex);
            }
        }

        public void OnRemoveSignatureClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || mail.Sent) return;
                RemoveFlag(mail, SECFLAG_SIGNED, "Sign");
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error("Sign", "Remove signature failed", ex);
            }
        }

        // ── Decrypt ─────────────────────────────────────────────────────────
        public void OnDecryptClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null)
                {
                    Logger.Warn("Decrypt", "Decrypt clicked but no active mail item found");
                    return;
                }

                Logger.Info("Decrypt",
                    $"Decrypt clicked — from: {mail.SenderEmailAddress}, " +
                    $"subject: {Truncate(mail.Subject, 30)}");

                Outlook.Attachment? smimeAttachment = null;
                for (int i = 1; i <= mail.Attachments.Count; i++)
                {
                    var att = mail.Attachments[i];
                    if (att.FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                    {
                        smimeAttachment = att;
                        break;
                    }
                }

                if (smimeAttachment == null)
                {
                    Logger.Info("Decrypt",
                        "No .p7m attachment found — message may already be decrypted by Outlook");
                    MessageBox.Show(
                        "No encrypted content found.\n\n" +
                        "Outlook automatically decrypts properly formatted S/MIME messages.\n" +
                        "This button decrypts Parcl at-rest encrypted messages and .p7m attachments.",
                        "Parcl — Decrypt",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var tempPath = Path.Combine(Path.GetTempPath(), "parcl-decrypt.p7m");
                smimeAttachment.SaveAsFile(tempPath);
                byte[] encryptedData;
                try
                {
                    encryptedData = File.ReadAllBytes(tempPath);
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }

                var result = SmimeHandler.Decrypt(encryptedData);
                if (result.Success && result.Content != null)
                {
                    bool isAtRest = smimeAttachment.FileName == "parcl-encrypted.p7m";

                    mail.Body = Encoding.UTF8.GetString(result.Content);

                    if (isAtRest)
                    {
                        for (int i = mail.Attachments.Count; i >= 1; i--)
                        {
                            if (mail.Attachments[i].FileName == "parcl-encrypted.p7m")
                                mail.Attachments[i].Delete();
                        }
                        mail.Save();
                    }

                    Logger.Info("Decrypt",
                        isAtRest ? "At-rest message decrypted and restored"
                                 : "Attached .p7m decrypted successfully");
                    MessageBox.Show("Message decrypted successfully.",
                        "Parcl — Decrypted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Logger.Error("Decrypt", $"Decryption failed: {result.ErrorMessage}");
                    MessageBox.Show($"Decryption failed: {result.ErrorMessage}",
                        "Parcl — Decrypt Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Decrypt", "Decryption action failed", ex);
                MessageBox.Show($"Decryption error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Certificate Exchange ────────────────────────────────────────────
        public void OnCertExchangeClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || mail.Sent)
                {
                    Logger.Warn("Exchange",
                        "No active compose item to attach certificate to");
                    return;
                }

                Logger.Info("Exchange", "Certificate exchange dialog opening");
                using (var dialog = new CertExchangeDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK &&
                        dialog.SelectedCertificate != null)
                    {
                        var certInfo = dialog.SelectedCertificate;
                        var format = dialog.AttachAsPem ? "PEM" : "DER";

                        Logger.Info("Exchange",
                            $"Certificate selected for exchange: {certInfo.Subject} " +
                            $"[{certInfo.Thumbprint.Substring(0, 8)}], format={format}");

                        var payload = CertExchange.PrepareExport(certInfo.Thumbprint);
                        var attachment = dialog.AttachAsPem
                            ? Encoding.ASCII.GetBytes(
                                CertExchange.FormatAsAttachment(payload))
                            : Convert.FromBase64String(payload.CertificateData);

                        var extension = dialog.AttachAsPem ? ".pem" : ".cer";
                        var tempFile = Path.Combine(
                            Path.GetTempPath(), $"parcl-cert{extension}");
                        File.WriteAllBytes(tempFile, attachment);

                        mail.Attachments.Add(tempFile,
                            Outlook.OlAttachmentType.olByValue,
                            Type.Missing,
                            $"My Certificate ({certInfo.Subject}){extension}");

                        try { File.Delete(tempFile); } catch { }

                        Logger.Info("Exchange", "Certificate attached to message");
                    }
                    else
                    {
                        Logger.Debug("Exchange",
                            "Certificate exchange cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Exchange", "Certificate exchange failed", ex);
                MessageBox.Show($"Certificate exchange error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Certificate Selector ────────────────────────────────────────────
        public void OnCertSelectorClick(IRibbonControl control)
        {
            try
            {
                Logger.Info("CertSel", "Certificate selector dialog opening");
                using (var dialog = new CertificateSelectorDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        Logger.Info("CertSel", "Certificate selection saved");
                        _ribbon?.Invalidate();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("CertSel", "Certificate selector failed", ex);
                MessageBox.Show($"Certificate selector error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── LDAP Lookup ─────────────────────────────────────────────────────
        public void OnLookupClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null)
                {
                    Logger.Warn("Lookup", "Lookup clicked but no active mail item found");
                    return;
                }

                Logger.Info("Lookup",
                    $"LDAP lookup triggered — recipients: {mail.To}");

                var directories = Settings.LdapDirectories
                    .Where(d => d.Enabled).ToList();
                if (directories.Count == 0)
                {
                    MessageBox.Show(
                        "No LDAP directories configured. Open Options to add one.",
                        "Parcl — Lookup",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int found = 0;
                for (int i = 1; i <= mail.Recipients.Count; i++)
                {
                    var addr = mail.Recipients[i].Address;
                    var certs = LdapLookup
                        .LookupAcrossDirectoriesAsync(addr, directories).Result;
                    found += certs.Count;

                    foreach (var certInfo in certs)
                    {
                        if (CertStore.FindByThumbprint(certInfo.Thumbprint) == null)
                        {
                            Logger.Debug("Lookup",
                                $"Certificate {certInfo.Thumbprint.Substring(0, 8)} " +
                                "found via LDAP but needs import");
                        }
                    }

                    CertCache.Add(addr, certs);
                }

                Logger.Info("Lookup",
                    $"LDAP lookup complete — {found} certificate(s) found " +
                    $"for {mail.Recipients.Count} recipient(s)");
                MessageBox.Show(
                    $"LDAP lookup complete.\n\n" +
                    $"Recipients searched: {mail.Recipients.Count}\n" +
                    $"Certificates found: {found}",
                    "Parcl — Lookup Results",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("Lookup", "LDAP lookup failed", ex);
                MessageBox.Show($"Lookup error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Dashboard / Options ─────────────────────────────────────────────
        public void OnDashboardToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                Logger.Info("UI", $"Dashboard toggle: {(pressed ? "show" : "hide")}");
                ToggleTaskPane();
            }
            catch (Exception ex)
            {
                Logger.Error("UI", "Dashboard toggle failed", ex);
            }
        }

        public bool GetDashboardPressed(IRibbonControl control)
        {
            return false;
        }

        public void OnOptionsClick(IRibbonControl control)
        {
            try
            {
                Logger.Info("Options", "Options dialog opening");
                using (var dialog = new OptionsDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                        Logger.Info("Options", "Settings saved by user");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Options", "Options dialog failed", ex);
                MessageBox.Show($"Options error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Core encrypt/sign logic ─────────────────────────────────────────

        /// <summary>
        /// Marks a compose message for encryption. The message stays editable.
        /// Actual encapsulation happens in Application_ItemSend right before the message leaves.
        /// </summary>
        private void EncryptOutgoing(Outlook.MailItem mail)
        {
            Logger.Info("Encrypt",
                $"Encrypt toggled ON — to: {mail.To}, subject: {Truncate(mail.Subject, 30)}");

            // Set a user property flag so ItemSend knows to encapsulate
            var flag = mail.UserProperties.Find("ParclEncrypt") ??
                       mail.UserProperties.Add("ParclEncrypt", Outlook.OlUserPropertyType.olYesNo, false);
            flag.Value = true;

            Logger.Info("Encrypt", "Message flagged for S/MIME encryption on send");
        }

        /// <summary>
        /// Marks a compose message for signing. Validated at toggle time,
        /// actual PR_SECURITY_FLAGS set at send time in Application_ItemSend.
        /// </summary>
        private void SignOutgoing(Outlook.MailItem mail)
        {
            var thumbprint = Settings.UserProfile.SigningCertThumbprint;

            if (string.IsNullOrEmpty(thumbprint))
            {
                Logger.Warn("Sign", "No signing certificate selected");
                MessageBox.Show(
                    "No signing certificate selected. " +
                    "Use the Certificate Selector to choose one.",
                    "Parcl — Sign",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var signingCert = CertStore.FindByThumbprint(thumbprint!);
            if (signingCert == null || !signingCert.HasPrivateKey)
            {
                Logger.Error("Sign",
                    "Signing certificate not found or missing private key");
                MessageBox.Show(
                    "The selected signing certificate was not found or " +
                    "does not have a private key.",
                    "Parcl — Sign Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Set user property flag — actual signing happens at send time
            var flag = mail.UserProperties.Find("ParclSign") ??
                       mail.UserProperties.Add("ParclSign", Outlook.OlUserPropertyType.olYesNo, false);
            flag.Value = true;

            Logger.Info("Sign",
                $"Message flagged for S/MIME signing with cert {thumbprint!.Substring(0, 8)}");
        }

        private void EncryptAtRest(Outlook.MailItem mail)
        {
            Logger.Info("Encrypt",
                $"Encrypt at rest — subject: {Truncate(mail.Subject, 30)}");

            var thumbprint = Settings.UserProfile.EncryptionCertThumbprint;
            if (string.IsNullOrEmpty(thumbprint))
            {
                MessageBox.Show(
                    "No encryption certificate selected.\n" +
                    "Use the Certificate Selector to choose one first.",
                    "Parcl — Encrypt",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var cert = CertStore.FindByThumbprint(thumbprint!);
            if (cert == null)
            {
                MessageBox.Show(
                    "Encryption certificate not found in your certificate store.",
                    "Parcl — Encrypt Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var bodyBytes = Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
            var recipientCerts = new X509Certificate2Collection { cert };
            var encrypted = SmimeHandler.Encrypt(bodyBytes, recipientCerts);

            var tempPath = Path.Combine(Path.GetTempPath(), "parcl-atrest.p7m");
            File.WriteAllBytes(tempPath, encrypted);

            mail.Body = "[Parcl] This message has been encrypted at rest.\n" +
                        "Open with Parcl Decrypt to read the original content.";

            for (int i = mail.Attachments.Count; i >= 1; i--)
            {
                if (mail.Attachments[i].FileName == "parcl-encrypted.p7m")
                    mail.Attachments[i].Delete();
            }

            mail.Attachments.Add(tempPath,
                Outlook.OlAttachmentType.olByValue,
                Type.Missing,
                "parcl-encrypted.p7m");

            try { File.Delete(tempPath); } catch { }

            mail.Save();
            Logger.Info("Encrypt", "Message encrypted at rest and saved");
        }

        private void RemoveFlag(Outlook.MailItem mail, int flag, string component)
        {
            if (flag == SECFLAG_ENCRYPTED)
            {
                var prop = mail.UserProperties.Find("ParclEncrypt");
                if (prop != null)
                    prop.Value = false;
                Logger.Info(component, "Encryption removed from message");
            }

            if (flag == SECFLAG_SIGNED)
            {
                var prop = mail.UserProperties.Find("ParclSign");
                if (prop != null)
                    prop.Value = false;
                Logger.Info(component, "Signing removed from message");
            }
        }

        // ── Cert resolution: Outlook contacts → GAL/Exchange → AddressEntry → cert stores ──

        private X509Certificate2? ResolveRecipientCert(string email, Outlook.Recipient recipient)
        {
            // 1. Try AddressEntry's X.509 certificate property (works for GAL, Exchange, and contacts)
            try
            {
                var addrEntry = recipient.AddressEntry;
                if (addrEntry != null)
                {
                    // PR_USER_X509_CERTIFICATE is a multi-valued binary property (PT_MV_BINARY = 0x1102)
                    try
                    {
                        var certValues = addrEntry.PropertyAccessor.GetProperty(PR_USER_X509_CERT);
                        if (certValues is object[] certArray)
                        {
                            foreach (var item in certArray)
                            {
                                if (item is byte[] certData && certData.Length > 0)
                                {
                                    try
                                    {
                                        var cert = new X509Certificate2(certData);
                                        if (cert.NotAfter > DateTime.UtcNow)
                                        {
                                            Logger.Debug("Encrypt",
                                                $"Certificate found via AddressEntry for {email}");
                                            return cert;
                                        }
                                    }
                                    catch { /* malformed cert data, try next */ }
                                }
                            }
                        }
                    }
                    catch { /* Property not available on this AddressEntry */ }

                    // 2. Try Exchange user object (GAL users)
                    try
                    {
                        var exchUser = addrEntry.GetExchangeUser();
                        if (exchUser != null)
                        {
                            try
                            {
                                var certValues = exchUser.PropertyAccessor.GetProperty(PR_USER_X509_CERT);
                                if (certValues is object[] exchCerts)
                                {
                                    foreach (var item in exchCerts)
                                    {
                                        if (item is byte[] certData && certData.Length > 0)
                                        {
                                            try
                                            {
                                                var cert = new X509Certificate2(certData);
                                                if (cert.NotAfter > DateTime.UtcNow)
                                                {
                                                    Logger.Debug("Encrypt",
                                                        $"Certificate found via Exchange GAL for {email}");
                                                    return cert;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            catch { }

                            Marshal.ReleaseComObject(exchUser);
                        }
                    }
                    catch { /* Not an Exchange user */ }

                    // 3. Try Outlook contact object
                    try
                    {
                        var contact = addrEntry.GetContact();
                        if (contact != null)
                        {
                            try
                            {
                                var certValues = contact.PropertyAccessor.GetProperty(PR_USER_X509_CERT);
                                if (certValues is object[] contactCerts)
                                {
                                    foreach (var item in contactCerts)
                                    {
                                        if (item is byte[] certData && certData.Length > 0)
                                        {
                                            try
                                            {
                                                var cert = new X509Certificate2(certData);
                                                if (cert.NotAfter > DateTime.UtcNow)
                                                {
                                                    Logger.Debug("Encrypt",
                                                        $"Certificate found in Outlook contact for {email}");
                                                    return cert;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            catch { }

                            Marshal.ReleaseComObject(contact);
                        }
                    }
                    catch { /* Contact lookup not available */ }
                }
            }
            catch { /* AddressEntry access failed */ }

            // 4. Windows certificate stores (AddressBook → My)
            var storeCert = CertStore.FindByEmail(email);
            if (storeCert != null)
            {
                Logger.Debug("Encrypt", $"Certificate found in Windows store for {email}");
                return storeCert;
            }

            Logger.Debug("Encrypt", $"No certificate found for {email} in any source");
            return null;
        }

        // ── SMTP address resolution ──────────────────────────────────────────

        /// <summary>
        /// Resolves the SMTP email address for a recipient. Exchange internal users
        /// return X500 addresses from recipient.Address — this resolves to the real SMTP.
        /// </summary>
        private static string ResolveSmtpAddress(Outlook.Recipient recipient)
        {
            try
            {
                var addrEntry = recipient.AddressEntry;
                if (addrEntry == null)
                    return recipient.Address;

                // If it's already SMTP, use it directly
                if (addrEntry.Type == "SMTP")
                    return recipient.Address;

                // Exchange user — get PrimarySmtpAddress
                if (addrEntry.Type == "EX")
                {
                    try
                    {
                        var exchUser = addrEntry.GetExchangeUser();
                        if (exchUser != null)
                        {
                            var smtp = exchUser.PrimarySmtpAddress;
                            Marshal.ReleaseComObject(exchUser);
                            if (!string.IsNullOrEmpty(smtp))
                                return smtp;
                        }
                    }
                    catch { }

                    // Fallback: read PR_SMTP_ADDRESS from the AddressEntry
                    try
                    {
                        const string PR_SMTP_ADDRESS =
                            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        var smtp = (string)addrEntry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
                        if (!string.IsNullOrEmpty(smtp))
                            return smtp;
                    }
                    catch { }
                }

                return recipient.Address;
            }
            catch
            {
                return recipient.Address;
            }
        }

        // ── Import Certificates (ribbon button — manual only) ─────────────

        public void OnImportCertificatesClick(IRibbonControl control)
        {
            try
            {
                Outlook.MailItem? mail = GetMailItem(control);
                if (mail == null || !mail.Sent)
                {
                    Logger.Warn("Import", "Import clicked but no received mail item found");
                    return;
                }

                int count = ImportCertificateAttachments(mail);
                if (count > 0)
                {
                    MessageBox.Show(
                        $"Imported {count} certificate(s) from this message.",
                        "Parcl — Import Certificates",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "No certificate attachments found in this message.",
                        "Parcl — Import Certificates",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Import", "Certificate import failed", ex);
                MessageBox.Show($"Import error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Scans a received message for certificate attachments (.cer, .pem, .crt, .der, .p7c)
        /// and imports them into the AddressBook store after prompting the user for consent.
        /// </summary>
        internal int ImportCertificateAttachments(Outlook.MailItem mail)
        {
            int imported = 0;
            var certExtensions = new[] { ".cer", ".pem", ".crt", ".der", ".p7c" };

            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var att = mail.Attachments[i];
                var ext = Path.GetExtension(att.FileName)?.ToLowerInvariant();
                if (ext == null || !certExtensions.Contains(ext))
                    continue;

                var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ext);
                try
                {
                    att.SaveAsFile(tempPath);
                    var rawData = File.ReadAllBytes(tempPath);

                    // Handle PEM: strip header/footer and decode base64
                    byte[] certData;
                    if (ext == ".pem")
                    {
                        var pem = Encoding.ASCII.GetString(rawData);
                        pem = pem
                            .Replace("-----BEGIN CERTIFICATE-----", "")
                            .Replace("-----END CERTIFICATE-----", "")
                            .Replace("\r", "").Replace("\n", "").Trim();
                        certData = Convert.FromBase64String(pem);
                    }
                    else
                    {
                        certData = rawData;
                    }

                    var cert = new X509Certificate2(certData);
                    if (cert.NotAfter <= DateTime.UtcNow)
                    {
                        Logger.Debug("Import", $"Skipping expired cert from {att.FileName}");
                        continue;
                    }

                    var info = Parcl.Core.Models.CertificateInfo.FromX509(cert);

                    // Prompt user before importing each certificate
                    var prompt = MessageBox.Show(
                        $"Import this certificate?\n\n" +
                        $"Subject: {info.Subject}\n" +
                        $"Issuer: {info.Issuer}\n" +
                        $"Thumbprint: {info.Thumbprint}\n" +
                        $"Valid: {cert.NotBefore:d} — {cert.NotAfter:d}",
                        "Parcl — Confirm Certificate Import",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (prompt != DialogResult.Yes)
                    {
                        Logger.Debug("Import", $"User declined import of {info.Subject}");
                        continue;
                    }

                    // Import to AddressBook store (Other People)
                    CertStore.PublishToAddressBook(cert);
                    imported++;

                    Logger.Info("Import",
                        $"Certificate imported from attachment: {info.Subject} " +
                        $"[{info.Thumbprint.Substring(0, 8)}] from {mail.SenderEmailAddress}");

                    // Cache it for the sender
                    if (!string.IsNullOrEmpty(mail.SenderEmailAddress))
                    {
                        CertCache.Add(mail.SenderEmailAddress,
                            new List<Parcl.Core.Models.CertificateInfo> { info });
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn("Import",
                        $"Failed to import cert from {att.FileName}: {ex.Message}");
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }

            return imported;
        }

        // ── Helpers ─────────────────────────────────────────────────────────

        private Outlook.MailItem? GetMailItem(IRibbonControl control)
        {
            if (control.Context is Outlook.Inspector inspector)
                return inspector.CurrentItem as Outlook.MailItem;

            if (control.Context is Outlook.Explorer explorer)
            {
                var selection = explorer.Selection;
                if (selection.Count > 0 && selection[1] is Outlook.MailItem selected)
                    return selected;
            }

            return null;
        }

        /// <summary>
        /// Safe version for getPressed callbacks — must never throw.
        /// </summary>
        private Outlook.MailItem? GetMailItemSafe(IRibbonControl control)
        {
            try { return GetMailItem(control); }
            catch { return null; }
        }

        private static int GetSecurityFlags(Outlook.PropertyAccessor pa)
        {
            try { return (int)pa.GetProperty(PR_SECURITY_FLAGS); }
            catch { return 0; }
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException(
                        $"Resource {resourceName} not found.");
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        private static string Truncate(string? s, int maxLen)
        {
            if (string.IsNullOrEmpty(s)) return "(empty)";
            return s!.Length <= maxLen ? s : s.Substring(0, maxLen) + "...";
        }
    }
}
