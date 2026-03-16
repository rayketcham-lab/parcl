using System;
using Parcl.Addin.TaskPane;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Parcl.Addin
{
    public partial class ThisAddIn
    {
        internal ParclSettings Settings { get; private set; } = null!;
        internal CertificateStore CertStore { get; private set; } = null!;
        internal SmimeHandler SmimeHandler { get; private set; } = null!;
        internal LdapCertLookup LdapLookup { get; private set; } = null!;
        internal CertificateCache CertCache { get; private set; } = null!;
        internal CertExchange CertExchange { get; private set; } = null!;
        internal ParclLogger Logger { get; private set; } = null!;

        private CustomTaskPane? _taskPane;
        private ParclTaskPaneHost? _taskPaneHost;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger = new ParclLogger();
            Logger.Info("AddIn", "Parcl add-in starting up");

            try
            {
                Settings = ParclSettings.Load();
                Logger.Debug("AddIn", $"Settings loaded — {Settings.LdapDirectories.Count} LDAP directories configured");

                CertStore = new CertificateStore();
                SmimeHandler = new SmimeHandler();
                LdapLookup = new LdapCertLookup(Logger);
                CertCache = new CertificateCache(
                    Settings.Cache.CacheExpirationHours,
                    Settings.Cache.MaxCacheEntries);
                CertExchange = new CertExchange(CertStore);

                Logger.Debug("AddIn", "Core services initialized");

                // Create the animated task pane
                _taskPaneHost = new ParclTaskPaneHost();
                _taskPane = CustomTaskPanes.Add(_taskPaneHost, "Parcl");
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 320;
                _taskPane.Visible = false;

                Logger.Info("AddIn", "Task pane created (docked right, 320px)");

                if (Settings.Behavior.AutoLookup == LookupTrigger.OnSend)
                {
                    Application.ItemSend += Application_ItemSend;
                    Logger.Debug("AddIn", "Auto-lookup on send enabled");
                }

                Logger.Info("AddIn", "Parcl add-in started successfully");
            }
            catch (Exception ex)
            {
                Logger.Error("AddIn", "Failed during startup", ex);
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            if (item is Outlook.MailItem mail)
            {
                var subjectPreview = mail.Subject != null
                    ? mail.Subject.Substring(0, Math.Min(30, mail.Subject.Length))
                    : "(no subject)";
                Logger.Info("Send", $"ItemSend intercepted — to: {mail.To}, subject: {subjectPreview}");

                try
                {
                    if (Settings.Crypto.AlwaysSign && !string.IsNullOrEmpty(Settings.UserProfile.SigningCertThumbprint))
                    {
                        Logger.Info("Send", "Auto-sign enabled — applying digital signature");
                        var signingCert = CertStore.FindByThumbprint(Settings.UserProfile.SigningCertThumbprint);
                        if (signingCert != null && signingCert.HasPrivateKey)
                        {
                            var bodyBytes = System.Text.Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                            var signed = SmimeHandler.Sign(bodyBytes, signingCert);
                            mail.PropertyAccessor.SetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x6E010102", signed);
                            Logger.Info("Send", "Digital signature applied successfully");
                        }
                        else
                        {
                            Logger.Warn("Send", "Signing certificate not found or has no private key — skipping auto-sign");
                        }
                    }

                    if (Settings.Crypto.AlwaysEncrypt)
                    {
                        Logger.Info("Send", "Auto-encrypt enabled — looking up recipient certificates");
                        var recipients = mail.Recipients;
                        var recipientCerts = new System.Security.Cryptography.X509Certificates.X509Certificate2Collection();
                        bool allResolved = true;

                        for (int i = 1; i <= recipients.Count; i++)
                        {
                            var recipientEmail = recipients[i].Address;
                            var cert = CertStore.FindByEmail(recipientEmail);
                            if (cert != null)
                            {
                                recipientCerts.Add(cert);
                            }
                            else
                            {
                                Logger.Warn("Send", $"No encryption certificate found for {recipientEmail}");
                                allResolved = false;
                            }
                        }

                        if (!allResolved && Settings.Behavior.PromptOnMissingCert)
                        {
                            var result = System.Windows.Forms.MessageBox.Show(
                                "Encryption certificates could not be found for all recipients. Send unencrypted?",
                                "Parcl — Missing Certificates",
                                System.Windows.Forms.MessageBoxButtons.YesNo,
                                System.Windows.Forms.MessageBoxIcon.Warning);
                            if (result == System.Windows.Forms.DialogResult.No)
                            {
                                cancel = true;
                                Logger.Info("Send", "User cancelled send due to missing recipient certificates");
                                return;
                            }
                        }

                        if (recipientCerts.Count > 0 && allResolved)
                        {
                            var bodyBytes = System.Text.Encoding.UTF8.GetBytes(mail.Body ?? string.Empty);
                            var encrypted = SmimeHandler.Encrypt(bodyBytes, recipientCerts);
                            mail.PropertyAccessor.SetProperty(
                                "http://schemas.microsoft.com/mapi/proptag/0x6E010102", encrypted);
                            Logger.Info("Send", $"Message encrypted for {recipientCerts.Count} recipient(s)");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Send", "Failed during auto-encrypt/sign", ex);
                    System.Windows.Forms.MessageBox.Show(
                        $"Parcl encountered an error during send processing:\n{ex.Message}",
                        "Parcl Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        internal void ToggleTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
                Logger.Debug("UI", $"Task pane toggled: {(_taskPane.Visible ? "visible" : "hidden")}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Info("AddIn", "Parcl add-in shutting down");
            Settings?.Save();
            CertStore?.Dispose();
            Logger?.Dispose();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ParclRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
