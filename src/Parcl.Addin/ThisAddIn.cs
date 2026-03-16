using System;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Settings = ParclSettings.Load();
            CertStore = new CertificateStore();
            SmimeHandler = new SmimeHandler();
            LdapLookup = new LdapCertLookup();
            CertCache = new CertificateCache(
                Settings.Cache.CacheExpirationHours,
                Settings.Cache.MaxCacheEntries);
            CertExchange = new CertExchange(CertStore);

            if (Settings.Behavior.AutoLookup == LookupTrigger.OnSend)
            {
                Application.ItemSend += Application_ItemSend;
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            if (item is Outlook.MailItem mail)
            {
                // Auto-lookup and encrypt on send if configured
                // Implementation will wire into the full send pipeline
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Settings.Save();
            CertStore?.Dispose();
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
