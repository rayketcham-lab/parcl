using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Parcl.Addin.Dialogs;
using Parcl.Core.Config;
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
                    MessageBox.Show(
                        "Encrypt functionality will resolve recipient certificates via LDAP and encrypt the message body.",
                        "Parcl - Encrypt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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
                    MessageBox.Show(
                        "Decrypt functionality will use your private key to decrypt S/MIME encrypted messages.",
                        "Parcl - Decrypt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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
                    var settings = ParclSettings.Load();
                    var thumbprint = settings.UserProfile.SigningCertThumbprint;

                    if (string.IsNullOrEmpty(thumbprint))
                    {
                        _logger.Warn("Sign", "Sign clicked but no signing certificate selected");
                        MessageBox.Show(
                            "No signing certificate selected. Use the Certificate Selector to choose one.",
                            "Parcl - Sign",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    _logger.Info("Sign", $"Sign clicked — using cert {thumbprint.Substring(0, 8)}, subject: {Truncate(mail.Subject, 30)}");
                    MessageBox.Show(
                        "Sign functionality will apply a digital signature using your selected signing certificate.",
                        "Parcl - Sign",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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
                _logger.Info("Exchange", "Certificate exchange dialog opening");
                using (var dialog = new CertExchangeDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK && dialog.SelectedCertificate != null)
                    {
                        _logger.Info("Exchange",
                            $"Certificate selected for exchange: {dialog.SelectedCertificate.Subject} " +
                            $"[{dialog.SelectedCertificate.Thumbprint.Substring(0, 8)}], format={(dialog.AttachAsPem ? "PEM" : "DER")}");
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
                    _logger.Info("Lookup", $"LDAP lookup triggered from ribbon — recipients: {mail.To}");
                    MessageBox.Show(
                        "Lookup will search configured LDAP directories for recipient certificates.",
                        "Parcl - Lookup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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
