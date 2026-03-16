using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Models;

namespace Parcl.Addin.Dialogs
{
    public class CertificateSelectorDialog : Form
    {
        private readonly CertificateStore _certStore;
        private readonly ParclSettings _settings;
        private ListView _signingListView = null!;
        private ListView _encryptionListView = null!;
        private Label _signingInfoLabel = null!;
        private Label _encryptionInfoLabel = null!;

        public CertificateSelectorDialog()
        {
            _certStore = new CertificateStore();
            _settings = ParclSettings.Load();
            InitializeComponents();
            LoadCertificates();
        }

        private void InitializeComponents()
        {
            Text = "Parcl - Certificate Selector";
            Size = new Size(700, 550);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(12, 4)
            };

            // Signing tab
            var signingTab = new TabPage("Signing Certificate");
            _signingListView = CreateCertListView();
            _signingInfoLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(8),
                Text = "Select the certificate to use for digitally signing emails."
            };
            signingTab.Controls.Add(_signingListView);
            signingTab.Controls.Add(_signingInfoLabel);

            // Encryption tab
            var encryptionTab = new TabPage("Encryption Certificate");
            _encryptionListView = CreateCertListView();
            _encryptionInfoLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(8),
                Text = "Select the certificate to use for decrypting received emails."
            };
            encryptionTab.Controls.Add(_encryptionListView);
            encryptionTab.Controls.Add(_encryptionInfoLabel);

            tabControl.TabPages.Add(signingTab);
            tabControl.TabPages.Add(encryptionTab);

            // Button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 45,
                Padding = new Padding(8)
            };

            var cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
            var okBtn = new Button { Text = "OK" };
            okBtn.Click += OkButton_Click;

            buttonPanel.Controls.Add(cancelBtn);
            buttonPanel.Controls.Add(okBtn);

            Controls.Add(tabControl);
            Controls.Add(buttonPanel);

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private ListView CreateCertListView()
        {
            var lv = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                GridLines = true
            };
            lv.Columns.Add("Subject", 200);
            lv.Columns.Add("Issuer", 150);
            lv.Columns.Add("Expires", 100);
            lv.Columns.Add("Thumbprint", 120);
            lv.Columns.Add("Key Usage", 100);
            lv.SelectedIndexChanged += CertList_SelectedIndexChanged;
            return lv;
        }

        private void LoadCertificates()
        {
            var signingCerts = _certStore.GetSigningCertificates();
            foreach (var cert in signingCerts)
            {
                var item = new ListViewItem(new[]
                {
                    cert.Subject,
                    cert.Issuer,
                    cert.NotAfter.ToString("yyyy-MM-dd"),
                    cert.Thumbprint.Substring(0, 16) + "...",
                    "Digital Signature"
                }) { Tag = cert };

                if (cert.Thumbprint == _settings.UserProfile.SigningCertThumbprint)
                    item.Selected = true;

                _signingListView.Items.Add(item);
            }

            var encryptionCerts = _certStore.GetEncryptionCertificates();
            foreach (var cert in encryptionCerts)
            {
                var item = new ListViewItem(new[]
                {
                    cert.Subject,
                    cert.Issuer,
                    cert.NotAfter.ToString("yyyy-MM-dd"),
                    cert.Thumbprint.Substring(0, 16) + "...",
                    "Key Encipherment"
                }) { Tag = cert };

                if (cert.Thumbprint == _settings.UserProfile.EncryptionCertThumbprint)
                    item.Selected = true;

                _encryptionListView.Items.Add(item);
            }
        }

        private void CertList_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (sender is ListView lv && lv.SelectedItems.Count > 0)
            {
                var cert = (CertificateInfo)lv.SelectedItems[0].Tag;
                var label = lv == _signingListView ? _signingInfoLabel : _encryptionInfoLabel;
                label.Text = $"Subject: {cert.Subject}\n" +
                             $"Serial: {cert.SerialNumber}\n" +
                             $"Valid: {cert.NotBefore:yyyy-MM-dd} to {cert.NotAfter:yyyy-MM-dd}";
            }
        }

        private void OkButton_Click(object? sender, EventArgs e)
        {
            if (_signingListView.SelectedItems.Count > 0)
            {
                var cert = (CertificateInfo)_signingListView.SelectedItems[0].Tag;
                _settings.UserProfile.SigningCertThumbprint = cert.Thumbprint;
            }

            if (_encryptionListView.SelectedItems.Count > 0)
            {
                var cert = (CertificateInfo)_encryptionListView.SelectedItems[0].Tag;
                _settings.UserProfile.EncryptionCertThumbprint = cert.Thumbprint;
            }

            _settings.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _certStore?.Dispose();
            base.Dispose(disposing);
        }
    }
}
