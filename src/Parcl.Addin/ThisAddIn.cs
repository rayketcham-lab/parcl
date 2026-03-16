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
                LdapLookup = new LdapCertLookup();
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
                Logger.Info("Send", $"ItemSend intercepted — to: {mail.To}, subject: {mail.Subject?.Substring(0, Math.Min(30, mail.Subject?.Length ?? 0))}");

                if (Settings.Crypto.AlwaysEncrypt)
                {
                    Logger.Info("Send", "Auto-encrypt enabled — checking recipient certificates");
                    // TODO: Wire into full encrypt pipeline
                }

                if (Settings.Crypto.AlwaysSign)
                {
                    Logger.Info("Send", "Auto-sign enabled — applying digital signature");
                    // TODO: Wire into full sign pipeline
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
