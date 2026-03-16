using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
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

        public void OnEncryptClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    Logger.Info("Encrypt",
                        $"Encrypt clicked — to: {mail.To}, subject: {Truncate(mail.Subject, 30)}");

                    var recipientCerts = new X509Certificate2Collection();
                    var missing = new List<string>();

                    for (int i = 1; i <= mail.Recipients.Count; i++)
                    {
                        var addr = mail.Recipients[i].Address;
                        var cert = CertStore.FindByEmail(addr);
                        if (cert != null)
                            recipientCerts.Add(cert);
                        else
                            missing.Add(addr);
                    }

                    if (missing.Count > 0)
                    {
                        Logger.Warn("Encrypt",
                            $"Missing certificates for: {string.Join(", ", missing)}");
                        MessageBox.Show(
                            $"Could not find encryption certificates for:\n" +
                            $"{string.Join("\n", missing)}\n\nUse LDAP Lookup to find them first.",
                            "Parcl — Missing Certificates",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var bodyBytes = Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                    var encrypted = SmimeHandler.Encrypt(bodyBytes, recipientCerts);

                    var tempPath = Path.Combine(Path.GetTempPath(), "parcl-smime.p7m");
                    File.WriteAllBytes(tempPath, encrypted);

                    mail.Body = "This is an S/MIME encrypted message.";
                    mail.Attachments.Add(tempPath,
                        Outlook.OlAttachmentType.olByValue,
                        Type.Missing,
                        "smime.p7m");

                    try { File.Delete(tempPath); } catch { }

                    Logger.Info("Encrypt",
                        $"Message encrypted for {recipientCerts.Count} recipient(s)");
                    MessageBox.Show(
                        $"Message encrypted for {recipientCerts.Count} recipient(s).",
                        "Parcl — Encrypted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Logger.Warn("Encrypt", "Encrypt clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Encrypt", "Encryption action failed", ex);
                MessageBox.Show($"Encryption error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDecryptClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
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
                        Logger.Warn("Decrypt",
                            "No S/MIME encrypted content found in this message");
                        MessageBox.Show(
                            "This message does not appear to contain S/MIME encrypted content.\n" +
                            "Expected a .p7m attachment.",
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
                        mail.Body = Encoding.UTF8.GetString(result.Content);
                        Logger.Info("Decrypt", "Message decrypted successfully");
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
                else
                {
                    Logger.Warn("Decrypt", "Decrypt clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Decrypt", "Decryption action failed", ex);
                MessageBox.Show($"Decryption error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSignClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    var thumbprint = Settings.UserProfile.SigningCertThumbprint;

                    if (string.IsNullOrEmpty(thumbprint))
                    {
                        Logger.Warn("Sign",
                            "Sign clicked but no signing certificate selected");
                        MessageBox.Show(
                            "No signing certificate selected. " +
                            "Use the Certificate Selector to choose one.",
                            "Parcl — Sign",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    Logger.Info("Sign",
                        $"Sign clicked — using cert {thumbprint.Substring(0, 8)}, " +
                        $"subject: {Truncate(mail.Subject, 30)}");

                    var signingCert = CertStore.FindByThumbprint(thumbprint);
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

                    var bodyBytes = Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                    var signed = SmimeHandler.Sign(bodyBytes, signingCert);

                    var tempPath = Path.Combine(Path.GetTempPath(), "parcl-smime.p7s");
                    File.WriteAllBytes(tempPath, signed);

                    mail.Attachments.Add(tempPath,
                        Outlook.OlAttachmentType.olByValue,
                        Type.Missing,
                        "smime.p7s");

                    try { File.Delete(tempPath); } catch { }

                    Logger.Info("Sign", "Digital signature applied successfully");
                    MessageBox.Show("Digital signature applied successfully.",
                        "Parcl — Signed",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Logger.Warn("Sign", "Sign clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Sign", "Signing action failed", ex);
                MessageBox.Show($"Signing error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnCertExchangeClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (!(inspector?.CurrentItem is Outlook.MailItem mail))
                {
                    Logger.Warn("Exchange",
                        "No active mail item to attach certificate to");
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
                        MessageBox.Show(
                            $"Your certificate has been attached to the message as {format}.",
                            "Parcl — Certificate Attached",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    else
                    {
                        Logger.Debug("CertSel",
                            "Certificate selector cancelled by user");
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

        public void OnLookupClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    Logger.Info("Lookup",
                        $"LDAP lookup triggered from ribbon — recipients: {mail.To}");

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
                else
                {
                    Logger.Warn("Lookup",
                        "Lookup clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Lookup", "LDAP lookup failed", ex);
                MessageBox.Show($"Lookup error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
                    {
                        Logger.Info("Options", "Settings saved by user");
                    }
                    else
                    {
                        Logger.Debug("Options", "Options cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Options", "Options dialog failed", ex);
                MessageBox.Show($"Options error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            return s.Length <= maxLen ? s : s.Substring(0, maxLen) + "...";
        }
    }
}
