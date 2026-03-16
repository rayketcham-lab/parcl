using System;
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
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
using Parcl.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parcl.Addin
{
    [ComVisible(true)]
    public class ParclRibbon : IRibbonExtensibility
    {
        private IRibbonUI? _ribbon;
        private readonly ParclLogger _logger = new ParclLogger();

        public string GetCustomUI(string ribbonID)
        {
            _logger.Debug("Ribbon", $"GetCustomUI called for ribbonID={ribbonID}");
            return GetResourceText("Parcl.Addin.ParclRibbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            _logger.Info("Ribbon", "Parcl ribbon loaded into Outlook");
        }

        public void OnEncryptClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    _logger.Info("Encrypt", $"Encrypt clicked — to: {mail.To}, subject: {Truncate(mail.Subject, 30)}");

                    var addin = Globals.ThisAddIn;
                    var recipientCerts = new X509Certificate2Collection();
                    var missing = new System.Collections.Generic.List<string>();

                    for (int i = 1; i <= mail.Recipients.Count; i++)
                    {
                        var addr = mail.Recipients[i].Address;
                        var cert = addin.CertStore.FindByEmail(addr);
                        if (cert != null)
                        {
                            recipientCerts.Add(cert);
                        }
                        else
                        {
                            missing.Add(addr);
                        }
                    }

                    if (missing.Count > 0)
                    {
                        _logger.Warn("Encrypt", $"Missing certificates for: {string.Join(", ", missing)}");
                        MessageBox.Show(
                            $"Could not find encryption certificates for:\n{string.Join("\n", missing)}\n\nUse LDAP Lookup to find them first.",
                            "Parcl — Missing Certificates",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var bodyBytes = Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                    var encrypted = addin.SmimeHandler.Encrypt(bodyBytes, recipientCerts);
                    mail.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x6E010102", encrypted);

                    _logger.Info("Encrypt", $"Message encrypted for {recipientCerts.Count} recipient(s)");
                    MessageBox.Show(
                        $"Message encrypted for {recipientCerts.Count} recipient(s).",
                        "Parcl — Encrypted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _logger.Warn("Encrypt", "Encrypt clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Encrypt", "Encryption action failed", ex);
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
                    _logger.Info("Decrypt", $"Decrypt clicked — from: {mail.SenderEmailAddress}, subject: {Truncate(mail.Subject, 30)}");

                    var addin = Globals.ThisAddIn;
                    byte[] encryptedData;
                    try
                    {
                        encryptedData = (byte[])mail.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x6E010102");
                    }
                    catch
                    {
                        _logger.Warn("Decrypt", "No S/MIME encrypted content found in this message");
                        MessageBox.Show("This message does not appear to contain S/MIME encrypted content.",
                            "Parcl — Decrypt", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    var result = addin.SmimeHandler.Decrypt(encryptedData);
                    if (result.Success && result.Content != null)
                    {
                        mail.Body = Encoding.UTF8.GetString(result.Content);
                        _logger.Info("Decrypt", "Message decrypted successfully");
                        MessageBox.Show("Message decrypted successfully.",
                            "Parcl — Decrypted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        _logger.Error("Decrypt", $"Decryption failed: {result.ErrorMessage}");
                        MessageBox.Show($"Decryption failed: {result.ErrorMessage}",
                            "Parcl — Decrypt Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    _logger.Warn("Decrypt", "Decrypt clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Decrypt", "Decryption action failed", ex);
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
                    var addin = Globals.ThisAddIn;
                    var thumbprint = addin.Settings.UserProfile.SigningCertThumbprint;

                    if (string.IsNullOrEmpty(thumbprint))
                    {
                        _logger.Warn("Sign", "Sign clicked but no signing certificate selected");
                        MessageBox.Show(
                            "No signing certificate selected. Use the Certificate Selector to choose one.",
                            "Parcl — Sign",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    _logger.Info("Sign", $"Sign clicked — using cert {thumbprint.Substring(0, 8)}, subject: {Truncate(mail.Subject, 30)}");

                    var signingCert = addin.CertStore.FindByThumbprint(thumbprint);
                    if (signingCert == null || !signingCert.HasPrivateKey)
                    {
                        _logger.Error("Sign", "Signing certificate not found or missing private key");
                        MessageBox.Show(
                            "The selected signing certificate was not found or does not have a private key.",
                            "Parcl — Sign Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var bodyBytes = Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                    var signed = addin.SmimeHandler.Sign(bodyBytes, signingCert);
                    mail.PropertyAccessor.SetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x6E010102", signed);

                    _logger.Info("Sign", "Digital signature applied successfully");
                    MessageBox.Show("Digital signature applied successfully.",
                        "Parcl — Signed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _logger.Warn("Sign", "Sign clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Sign", "Signing action failed", ex);
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
                    _logger.Warn("Exchange", "No active mail item to attach certificate to");
                    return;
                }

                _logger.Info("Exchange", "Certificate exchange dialog opening");
                using (var dialog = new CertExchangeDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.SelectedCertificate != null)
                    {
                        var addin = Globals.ThisAddIn;
                        var certInfo = dialog.SelectedCertificate;
                        var format = dialog.AttachAsPem ? "PEM" : "DER";

                        _logger.Info("Exchange",
                            $"Certificate selected for exchange: {certInfo.Subject} " +
                            $"[{certInfo.Thumbprint.Substring(0, 8)}], format={format}");

                        var payload = addin.CertExchange.PrepareExport(certInfo.Thumbprint);
                        var attachment = dialog.AttachAsPem
                            ? Encoding.ASCII.GetBytes(addin.CertExchange.FormatAsAttachment(payload))
                            : Convert.FromBase64String(payload.CertificateData);

                        var extension = dialog.AttachAsPem ? ".pem" : ".cer";
                        var tempFile = Path.Combine(Path.GetTempPath(), $"parcl-cert{extension}");
                        File.WriteAllBytes(tempFile, attachment);

                        mail.Attachments.Add(tempFile, Outlook.OlAttachmentType.olByValue,
                            Type.Missing, $"My Certificate ({certInfo.Subject}){extension}");

                        try { File.Delete(tempFile); } catch { }

                        _logger.Info("Exchange", "Certificate attached to message");
                        MessageBox.Show($"Your certificate has been attached to the message as {format}.",
                            "Parcl — Certificate Attached", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        _logger.Debug("Exchange", "Certificate exchange cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Exchange", "Certificate exchange failed", ex);
                MessageBox.Show($"Certificate exchange error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnCertSelectorClick(IRibbonControl control)
        {
            try
            {
                _logger.Info("CertSel", "Certificate selector dialog opening");
                using (var dialog = new CertificateSelectorDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        _logger.Info("CertSel", "Certificate selection saved");
                        _ribbon?.Invalidate();
                    }
                    else
                    {
                        _logger.Debug("CertSel", "Certificate selector cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("CertSel", "Certificate selector failed", ex);
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
                    var addin = Globals.ThisAddIn;
                    _logger.Info("Lookup", $"LDAP lookup triggered from ribbon — recipients: {mail.To}");

                    var directories = addin.Settings.LdapDirectories.Where(d => d.Enabled).ToList();
                    if (directories.Count == 0)
                    {
                        MessageBox.Show("No LDAP directories configured. Open Options to add one.",
                            "Parcl — Lookup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    int found = 0;
                    int imported = 0;
                    for (int i = 1; i <= mail.Recipients.Count; i++)
                    {
                        var addr = mail.Recipients[i].Address;
                        var certs = addin.LdapLookup.LookupAcrossDirectoriesAsync(addr, directories).Result;
                        found += certs.Count;

                        foreach (var certInfo in certs)
                        {
                            if (addin.CertStore.FindByThumbprint(certInfo.Thumbprint) == null)
                            {
                                var rawCert = addin.CertStore.ExportPublicCertificate(certInfo.Thumbprint);
                                if (rawCert == null)
                                {
                                    _logger.Debug("Lookup", $"Certificate {certInfo.Thumbprint.Substring(0, 8)} found via LDAP but needs import");
                                }
                                imported++;
                            }
                        }

                        addin.CertCache.Add(addr, certs);
                    }

                    _logger.Info("Lookup", $"LDAP lookup complete — {found} certificate(s) found for {mail.Recipients.Count} recipient(s)");
                    MessageBox.Show(
                        $"LDAP lookup complete.\n\nRecipients searched: {mail.Recipients.Count}\nCertificates found: {found}",
                        "Parcl — Lookup Results",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _logger.Warn("Lookup", "Lookup clicked but no active mail item found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Lookup", "LDAP lookup failed", ex);
                MessageBox.Show($"Lookup error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDashboardToggle(IRibbonControl control, bool pressed)
        {
            try
            {
                _logger.Info("UI", $"Dashboard toggle: {(pressed ? "show" : "hide")}");
                Globals.ThisAddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                _logger.Error("UI", "Dashboard toggle failed", ex);
            }
        }

        public bool GetDashboardPressed(IRibbonControl control)
        {
            return false; // Task pane starts hidden
        }

        public void OnOptionsClick(IRibbonControl control)
        {
            try
            {
                _logger.Info("Options", "Options dialog opening");
                using (var dialog = new OptionsDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        _logger.Info("Options", "Settings saved by user");
                    }
                    else
                    {
                        _logger.Debug("Options", "Options cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Options", "Options dialog failed", ex);
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
                    throw new InvalidOperationException($"Resource {resourceName} not found.");
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
