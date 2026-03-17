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
using Parcl.Core.Config;
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

        /// <summary>
        /// Controls Parcl tab visibility. Only visible in Mail view, not Calendar/People/Tasks.
        /// </summary>
        public bool GetParclTabVisible(IRibbonControl control)
        {
            try
            {
                if (control.Context is Outlook.Explorer explorer)
                {
                    var folder = explorer.CurrentFolder;
                    if (folder != null)
                    {
                        // Show Parcl tab only in mail-related folders
                        var folderType = folder.DefaultItemType;
                        return folderType == Outlook.OlItemType.olMailItem;
                    }
                }

                // Always show in inspector windows (compose/read)
                if (control.Context is Outlook.Inspector)
                    return true;
            }
            catch { }

            return false;
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
                MessageBox.Show(
                    $"Encryption toggle failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Verify your certificates in Parcl > Select Certificates and try again.",
                    "Parcl — Encrypt Error",
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

                // Check if message has a .p7m attachment (Parcl-encrypted)
                try
                {
                    for (int i = 1; i <= mail.Attachments.Count; i++)
                    {
                        if (mail.Attachments[i].FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
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
                MessageBox.Show(
                    $"Encryption failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Verify your certificates in Parcl > Select Certificates and try again.",
                    "Parcl — Encrypt Error",
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
                MessageBox.Show(
                    $"Signing failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Verify your signing certificate in Parcl > Select Certificates and try again.",
                    "Parcl — Sign Error",
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
                MessageBox.Show(
                    $"Signing failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Verify your signing certificate in Parcl > Select Certificates and try again.",
                    "Parcl — Sign Error",
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

                string senderAddr;
                try { senderAddr = mail.SenderEmailAddress ?? "(unknown)"; }
                catch { senderAddr = "(unavailable)"; }

                string subjectPreview;
                try { subjectPreview = Truncate(mail.Subject, 30); }
                catch { subjectPreview = "(unavailable)"; }

                Logger.Info("Decrypt",
                    $"Decrypt clicked — from: {ParclLogger.SanitizeEmail(senderAddr)}, subject: {subjectPreview}");

                Outlook.Attachment? smimeAttachment = null;
                int smimeAttachmentIndex = -1;
                for (int i = 1; i <= mail.Attachments.Count; i++)
                {
                    var att = mail.Attachments[i];
                    if (att.FileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase))
                    {
                        smimeAttachment = att;
                        smimeAttachmentIndex = i;
                        break;
                    }
                }

                if (smimeAttachment == null)
                {
                    // Check if Outlook already decrypted this message natively
                    bool wasEncrypted = false;
                    try
                    {
                        var secFlags = (int)mail.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS);
                        wasEncrypted = (secFlags & SECFLAG_ENCRYPTED) != 0;
                    }
                    catch { }

                    if (wasEncrypted && mail.Body?.Length > 0)
                    {
                        Logger.Info("Decrypt", "Message already decrypted by Outlook natively");
                        MessageBox.Show(
                            "This message is already decrypted.\n\n" +
                            "Outlook decrypted it automatically using your certificate.\n" +
                            "The message content is visible in the reading pane.",
                            "Parcl — Already Decrypted",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        Logger.Info("Decrypt", "No encrypted content found");
                        MessageBox.Show(
                            "This message is not encrypted.\n\n" +
                            "Use the Encrypt button when composing to encrypt outgoing messages.",
                            "Parcl — Not Encrypted",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return;
                }

                bool isAtRest = smimeAttachment.FileName == "parcl-encrypted.p7m";

                var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".p7m");
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

                // ── Step 1: Decrypt the CMS envelope ──
                var result = SmimeHandler.Decrypt(encryptedData);
                if (!result.Success || result.Content == null)
                {
                    Logger.Error("Decrypt", $"Decryption failed: {result.ErrorMessage}");
                    MessageBox.Show(
                        "Decryption failed — could not unlock this message.\n\n" +
                        "Why: This message was likely encrypted for a different certificate than the one currently installed on your machine.\n\n" +
                        "Fix: Go to Parcl > Select Certificates and verify your encryption certificate matches the one the sender used. " +
                        "If you recently changed certificates, ask the sender to re-encrypt using your current certificate.",
                        "Parcl — Decrypt Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                byte[] mimeBytes = result.Content;
                Logger.Debug("Decrypt", $"CMS envelope decrypted: {mimeBytes.Length} bytes");

                // ── Step 2: Unwrap SignedCms if present (sign-then-encrypt) ──
                bool wasSigned = false;
                string? signerInfo = null;
                try
                {
                    var signedCms = new System.Security.Cryptography.Pkcs.SignedCms();
                    signedCms.Decode(mimeBytes);
                    signedCms.CheckSignature(verifySignatureOnly: false);
                    mimeBytes = signedCms.ContentInfo.Content;
                    wasSigned = true;

                    if (signedCms.SignerInfos.Count > 0 && signedCms.SignerInfos[0].Certificate != null)
                    {
                        var sigCert = signedCms.SignerInfos[0].Certificate!;
                        signerInfo = sigCert.Subject;
                        Logger.Info("Decrypt",
                            $"Signature verified — signer: {sigCert.Subject}, " +
                            $"thumbprint: {sigCert.Thumbprint.Substring(0, 8)}");
                    }
                }
                catch (System.Security.Cryptography.CryptographicException)
                {
                    // Not a SignedCms — content is raw MIME, which is fine (encrypt-only)
                    Logger.Debug("Decrypt", "Content is not signed (encrypt-only)");
                }

                // ── Step 3: Parse the MIME content to extract actual body ──
                string mimeText = Encoding.UTF8.GetString(mimeBytes);

                // Extract protected headers (RFC 7508)
                var protectedHeaders = Parcl.Core.Crypto.MimeBuilder.ExtractProtectedHeaders(mimeText);

                // Extract the actual body (HTML or plain text) and attachments
                var extracted = Parcl.Core.Crypto.MimeBuilder.ExtractBody(mimeText);

                if (extracted.HasContent)
                {
                    // Restore the actual message body
                    if (!string.IsNullOrEmpty(extracted.HtmlBody))
                    {
                        mail.HTMLBody = extracted.HtmlBody;
                        Logger.Debug("Decrypt", "HTML body restored from MIME");
                    }
                    else if (!string.IsNullOrEmpty(extracted.TextBody))
                    {
                        mail.Body = extracted.TextBody;
                        Logger.Debug("Decrypt", "Plain text body restored from MIME");
                    }
                }
                else
                {
                    // Fallback: set as plain text (shouldn't happen with well-formed MIME)
                    mail.Body = mimeText;
                    Logger.Warn("Decrypt", "Could not parse MIME structure — set raw content as body");
                }

                // Restore protected subject
                if (protectedHeaders != null && !string.IsNullOrEmpty(protectedHeaders.Subject))
                {
                    mail.Subject = protectedHeaders.Subject;
                    Logger.Info("Decrypt",
                        $"Protected subject restored: {Truncate(protectedHeaders.Subject, 30)}");
                }

                // Remove the .p7m attachment
                for (int i = mail.Attachments.Count; i >= 1; i--)
                {
                    var fn = mail.Attachments[i].FileName;
                    if (fn.Equals("smime.p7m", StringComparison.OrdinalIgnoreCase) ||
                        fn.Equals("parcl-encrypted.p7m", StringComparison.OrdinalIgnoreCase))
                    {
                        mail.Attachments[i].Delete();
                    }
                }

                // Re-add any attachments that were inside the encrypted envelope
                foreach (var att in extracted.Attachments)
                {
                    var attTempPath = Path.Combine(
                        Path.GetTempPath(), Path.GetRandomFileName() + "_" + att.FileName);
                    try
                    {
                        File.WriteAllBytes(attTempPath, att.Data);
                        mail.Attachments.Add(attTempPath,
                            Outlook.OlAttachmentType.olByValue,
                            Type.Missing, att.FileName);
                        Logger.Debug("Decrypt", $"Attachment restored: {att.FileName}");
                    }
                    finally
                    {
                        try { File.Delete(attTempPath); } catch { }
                    }
                }

                // Reset message class back to normal
                try
                {
                    const string PR_MESSAGE_CLASS = "http://schemas.microsoft.com/mapi/proptag/0x001A001E";
                    mail.PropertyAccessor.SetProperty(PR_MESSAGE_CLASS, "IPM.Note");
                }
                catch { }

                mail.Save();

                // Build result message
                var status = new StringBuilder("Message decrypted successfully.");
                if (wasSigned)
                    status.Append($"\n\nSignature verified: {signerInfo ?? "Unknown signer"}");
                if (protectedHeaders?.Subject != null)
                    status.Append($"\nOriginal subject restored.");
                if (extracted.Attachments.Count > 0)
                    status.Append($"\n{extracted.Attachments.Count} attachment(s) restored.");

                Logger.Info("Decrypt",
                    $"Message fully decrypted and restored" +
                    (wasSigned ? " (signed)" : "") +
                    $" — {extracted.Attachments.Count} attachment(s)");

                MessageBox.Show(status.ToString(),
                    "Parcl — Decrypted",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("Decrypt", "Decryption action failed", ex);
                MessageBox.Show(
                    $"Decryption failed unexpectedly.\n\n" +
                    $"Why: {ex.Message}\n\n" +
                    "Fix: Ensure the message is a valid Parcl-encrypted message and that your encryption certificate is installed. " +
                    "Go to Parcl > Select Certificates to verify.",
                    "Parcl — Decrypt Error",
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
                            Path.GetTempPath(), Path.GetRandomFileName() + extension);
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
                MessageBox.Show(
                    $"Certificate exchange failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Ensure you have a valid certificate selected in Parcl > Select Certificates and try again.",
                    "Parcl — Exchange Error",
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
                MessageBox.Show(
                    $"Certificate selector could not open.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Ensure your Windows certificate store is accessible and try again.",
                    "Parcl — Certificate Selector Error",
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
                    $"LDAP lookup triggered — recipients: {ParclLogger.SanitizeEmail(mail.To)}");

                var directories = Settings.LdapDirectories
                    .Where(d => d.Enabled).ToList();
                if (directories.Count == 0)
                {
                    MessageBox.Show(
                        "No LDAP directories configured.\n\n" +
                        "Why: LDAP lookup requires at least one directory server to search for recipient certificates.\n\n" +
                        "Fix: Click Parcl > Options on the ribbon and add an LDAP directory under the Directories tab.",
                        "Parcl — Lookup",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int found = 0;
                Outlook.Recipients? recipients = null;
                try
                {
                    recipients = mail.Recipients;
                    for (int i = 1; i <= recipients.Count; i++)
                    {
                        Outlook.Recipient? rcpt = null;
                        try
                        {
                            rcpt = recipients[i];
                            var addr = rcpt.Address;
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
                        finally
                        {
                            if (rcpt != null) Marshal.ReleaseComObject(rcpt);
                        }
                    }
                }
                finally
                {
                    if (recipients != null) Marshal.ReleaseComObject(recipients);
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
                MessageBox.Show(
                    $"LDAP certificate lookup failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Check your LDAP directory settings in Parcl > Options and verify the server is reachable.",
                    "Parcl — Lookup Error",
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

        public void OnAboutClick(IRibbonControl control)
        {
            try
            {
                using (var dialog = new Dialogs.AboutDialog())
                    dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Error("About", "About dialog failed", ex);
            }
        }

        public void OnCertContactsClick(IRibbonControl control)
        {
            try
            {
                Logger.Info("CertContacts", "Certificate contacts dialog opening");
                using (var dialog = new Dialogs.CertContactsDialog())
                    dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Error("CertContacts", "Certificate contacts dialog failed", ex);
            }
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
                MessageBox.Show(
                    $"Options dialog could not open.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Try closing and reopening Outlook. If the issue persists, check %APPDATA%\\Parcl\\ for corrupted settings.",
                    "Parcl — Options Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Core encrypt/sign logic ─────────────────────────────────────────

        /// <summary>
        /// Marks a compose message for encryption. The message stays editable.
        /// Actual encapsulation happens in Application_ItemSend right before the message leaves.
        /// Auto-detects the user's encryption certificate from the personal store if not configured.
        /// Falls back to all personal-store certs with private keys if no key-encipherment cert found.
        /// Opens the cert selector when the cert is unavailable or when multiple candidates exist.
        /// </summary>
        private void EncryptOutgoing(Outlook.MailItem mail)
        {
            Logger.Info("Encrypt",
                $"Encrypt toggled ON — to: {ParclLogger.SanitizeEmail(mail.To)}, subject: {Truncate(mail.Subject, 30)}");

            // Always read from disk so we pick up thumbprints saved by the selector dialog
            Settings = Parcl.Core.Config.ParclSettings.Load();

            var thumbprint = Settings.UserProfile.EncryptionCertThumbprint;
            if (string.IsNullOrEmpty(thumbprint))
            {
                Logger.Info("Encrypt", "No encryption cert configured — attempting auto-detect from personal store");

                // Try key-encipherment certs first, then fall back to any cert with a private key.
                // Many valid email certs omit the KeyUsage extension entirely.
                var candidates = CertStore.GetEncryptionCertificates()
                    .Where(c => c.HasPrivateKey && c.NotAfter > DateTime.UtcNow)
                    .OrderByDescending(c => c.NotAfter)
                    .ToList();

                if (candidates.Count == 0)
                {
                    candidates = CertStore.GetAllCertificates()
                        .Where(c => c.HasPrivateKey && c.NotAfter > DateTime.UtcNow)
                        .OrderByDescending(c => c.NotAfter)
                        .ToList();
                    Logger.Info("Encrypt",
                        $"Fell back to all personal-store certs with private keys: {candidates.Count} found");
                }

                if (candidates.Count == 1)
                {
                    // Exactly one cert — auto-select silently
                    Settings.UserProfile.EncryptionCertThumbprint = candidates[0].Thumbprint;
                    Settings.Save();
                    Logger.Info("Encrypt",
                        $"Auto-selected only available cert: {candidates[0].Subject}, thumbprint: {candidates[0].Thumbprint.Substring(0, 8)}");
                }
                else
                {
                    // Zero or multiple — open the selector so the user can choose
                    Logger.Info("Encrypt",
                        candidates.Count == 0
                            ? "No personal-store certs found — opening selector"
                            : $"{candidates.Count} candidates — opening selector for user choice");
                    using (var dlg = new Dialogs.CertificateSelectorDialog())
                    {
                        if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        {
                            Logger.Info("Encrypt", "Cert selection cancelled — encryption not flagged");
                            return;
                        }
                    }
                    // Reload settings saved by the dialog
                    Settings = Parcl.Core.Config.ParclSettings.Load();
                    if (string.IsNullOrEmpty(Settings.UserProfile.EncryptionCertThumbprint))
                    {
                        Logger.Warn("Encrypt", "No encryption cert selected in dialog — encryption not flagged");
                        return;
                    }
                }
            }
            else
            {
                // Thumbprint is set — verify the cert is still in the store and not expired
                var cert = CertStore.FindByThumbprint(thumbprint!);
                if (cert == null || cert.NotAfter <= DateTime.UtcNow)
                {
                    Logger.Warn("Encrypt",
                        cert == null ? "Configured encryption cert not found — opening selector" : "Configured encryption cert is expired — opening selector");
                    using (var dlg = new Dialogs.CertificateSelectorDialog())
                    {
                        if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                            return;
                    }
                    Settings = Parcl.Core.Config.ParclSettings.Load();
                    if (string.IsNullOrEmpty(Settings.UserProfile.EncryptionCertThumbprint))
                        return;
                }
            }

            // Set a user property flag so ItemSend knows to encapsulate
            var flag = mail.UserProperties.Find("ParclEncrypt") ??
                       mail.UserProperties.Add("ParclEncrypt", Outlook.OlUserPropertyType.olYesNo, false);
            flag.Value = true;

            Logger.Info("Encrypt",
                $"Message flagged for S/MIME encryption on send (cert: {Settings.UserProfile.EncryptionCertThumbprint!.Substring(0, 8)})");
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
                    "No signing certificate selected.\n\n" +
                    "Why: You must select a personal certificate with a private key before you can sign messages.\n\n" +
                    "Fix: Click Parcl > Select Certificates on the ribbon and choose a signing certificate.",
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
                    "The selected signing certificate is unavailable.\n\n" +
                    "Why: The certificate was not found in your Windows certificate store, or it does not have a private key.\n\n" +
                    "Fix: Go to Parcl > Select Certificates and choose a valid signing certificate that has a private key.",
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

            // Check if this message is already S/MIME encrypted (has smime.p7m or parcl-encrypted.p7m)
            for (int i = 1; i <= mail.Attachments.Count; i++)
            {
                var fn = mail.Attachments[i].FileName;
                if (fn.Equals("smime.p7m", StringComparison.OrdinalIgnoreCase) ||
                    fn.Equals("parcl-encrypted.p7m", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Info("Encrypt", "Message is already encrypted — skipping encrypt-at-rest");
                    MessageBox.Show(
                        "This message is already encrypted.\n\n" +
                        "Use Decrypt first if you want to re-encrypt it at rest.",
                        "Parcl — Already Encrypted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            var thumbprint = Settings.UserProfile.EncryptionCertThumbprint;
            if (string.IsNullOrEmpty(thumbprint))
            {
                MessageBox.Show(
                    "No encryption certificate selected.\n\n" +
                    "Why: An encryption certificate is required to encrypt messages at rest.\n\n" +
                    "Fix: Click Parcl > Select Certificates on the ribbon and choose your encryption certificate.",
                    "Parcl — Encrypt",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var cert = CertStore.FindByThumbprint(thumbprint!);
            if (cert == null)
            {
                MessageBox.Show(
                    "Encryption certificate not found.\n\n" +
                    "Why: The previously selected encryption certificate is no longer in your Windows certificate store.\n\n" +
                    "Fix: Go to Parcl > Select Certificates and choose a valid encryption certificate.",
                    "Parcl — Encrypt Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var body = mail.Body ?? string.Empty;
            if (string.IsNullOrWhiteSpace(body))
            {
                Logger.Warn("Encrypt", "Cannot encrypt at rest — message body is empty");
                MessageBox.Show(
                    "Cannot encrypt an empty message.\n\n" +
                    "The message body has no content to encrypt.",
                    "Parcl — Encrypt",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var bodyBytes = Encoding.UTF8.GetBytes(body);
            var recipientCerts = new X509Certificate2Collection { cert };
            var encrypted = SmimeHandler.Encrypt(bodyBytes, recipientCerts);

            var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".p7m");
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
                Logger.Info(component, "Encryption flag removed");
            }

            if (flag == SECFLAG_SIGNED)
            {
                var prop = mail.UserProperties.Find("ParclSign");
                if (prop != null)
                    prop.Value = false;
                Logger.Info(component, "Signing flag removed");
            }

            // Clear PR_SECURITY_FLAGS completely — remove the specific bit
            try
            {
                var pa = mail.PropertyAccessor;
                int flags;
                try { flags = (int)pa.GetProperty(PR_SECURITY_FLAGS); }
                catch { flags = 0; }

                // Clear the requested flag. Also clear opaque sign (0x20) when clearing sign.
                int clearMask = flag;
                if ((flag & SECFLAG_SIGNED) != 0)
                    clearMask |= 0x20; // opaque sign flag

                int newFlags = flags & ~clearMask;
                pa.SetProperty(PR_SECURITY_FLAGS, newFlags);
                Logger.Debug(component, $"PR_SECURITY_FLAGS: 0x{flags:X} -> 0x{newFlags:X}");

                // If all security flags are now 0, reset message class to IPM.Note
                if (newFlags == 0)
                {
                    try
                    {
                        if (mail.MessageClass != "IPM.Note")
                        {
                            mail.MessageClass = "IPM.Note";
                            Logger.Debug(component, "MessageClass reset to IPM.Note");
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug(component, $"Failed to clear PR_SECURITY_FLAGS: {ex.Message}");
            }

            mail.Save();
            _ribbon?.Invalidate();
            Logger.Info(component, $"Security flag 0x{flag:X2} cleared, message saved");
        }

        // ── Cert resolution: Outlook contacts → GAL/Exchange → AddressEntry → cert stores ──

        private X509Certificate2? ResolveRecipientCert(string email, Outlook.Recipient recipient)
        {
            // 1. Try AddressEntry's X.509 certificate property (works for GAL, Exchange, and contacts)
            Outlook.AddressEntry? addrEntry = null;
            try
            {
                addrEntry = recipient.AddressEntry;
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
                    Outlook.ExchangeUser? exchUser = null;
                    try
                    {
                        exchUser = addrEntry.GetExchangeUser();
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
                        }
                    }
                    catch { /* Not an Exchange user */ }
                    finally
                    {
                        if (exchUser != null) Marshal.ReleaseComObject(exchUser);
                    }

                    // 3. Try Outlook contact object
                    Outlook.ContactItem? contact = null;
                    try
                    {
                        contact = addrEntry.GetContact();
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
                        }
                    }
                    catch { /* Contact lookup not available */ }
                    finally
                    {
                        if (contact != null) Marshal.ReleaseComObject(contact);
                    }
                }
            }
            catch { /* AddressEntry access failed */ }
            finally
            {
                if (addrEntry != null) Marshal.ReleaseComObject(addrEntry);
            }

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
                MessageBox.Show(
                    $"Certificate import failed.\n\nWhy: {ex.Message}\n\n" +
                    "Fix: Verify the message contains valid certificate attachments (.cer, .pem, .crt) and try again.",
                    "Parcl — Import Error",
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

                    string importSender;
                    try { importSender = mail.SenderEmailAddress ?? ""; }
                    catch { importSender = ""; }

                    Logger.Info("Import",
                        $"Certificate imported from attachment: {info.Subject} " +
                        $"[{info.Thumbprint.Substring(0, 8)}] from {ParclLogger.SanitizeEmail(importSender)}");

                    // Cache it for the sender
                    if (!string.IsNullOrEmpty(importSender))
                    {
                        CertCache.Add(importSender,
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
