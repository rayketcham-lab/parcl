using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Parcl.Addin.Dialogs;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parcl.Addin
{
    [ComVisible(true)]
    public class ParclRibbon : IRibbonExtensibility
    {
        private IRibbonUI? _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Parcl.Addin.ParclRibbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnEncryptClick(IRibbonControl control)
        {
            try
            {
                var inspector = control.Context as Outlook.Inspector;
                if (inspector?.CurrentItem is Outlook.MailItem mail)
                {
                    // TODO: Resolve recipient certs, encrypt body
                    MessageBox.Show(
                        "Encrypt functionality will resolve recipient certificates via LDAP and encrypt the message body.",
                        "Parcl - Encrypt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
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
                    // TODO: Detect S/MIME envelope, decrypt with user's private key
                    MessageBox.Show(
                        "Decrypt functionality will use your private key to decrypt S/MIME encrypted messages.",
                        "Parcl - Decrypt",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
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
                    // TODO: Sign message body with selected signing cert
                    MessageBox.Show(
                        "Sign functionality will apply a digital signature using your selected signing certificate.",
                        "Parcl - Sign",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Signing error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnCertExchangeClick(IRibbonControl control)
        {
            try
            {
                using (var dialog = new CertExchangeDialog())
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Certificate exchange error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnCertSelectorClick(IRibbonControl control)
        {
            try
            {
                using (var dialog = new CertificateSelectorDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // Selections saved within dialog
                        _ribbon?.Invalidate();
                    }
                }
            }
            catch (Exception ex)
            {
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
                    // TODO: Extract recipients, lookup certs via LDAP
                    MessageBox.Show(
                        "Lookup will search configured LDAP directories for recipient certificates.",
                        "Parcl - Lookup",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lookup error: {ex.Message}", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnOptionsClick(IRibbonControl control)
        {
            try
            {
                using (var dialog = new OptionsDialog())
                {
                    dialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
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
    }
}
