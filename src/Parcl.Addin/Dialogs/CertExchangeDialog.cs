using System;
using System.Drawing;
using System.Windows.Forms;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Models;

namespace Parcl.Addin.Dialogs
{
    public class CertExchangeDialog : Form
    {
        private readonly CertificateStore _certStore;
        private readonly ParclSettings _settings;
        private ListView _certListView = null!;
        private Label _infoLabel = null!;
        private RadioButton _rbAttachCert = null!;
        private RadioButton _rbAttachPem = null!;

        public CertificateInfo? SelectedCertificate { get; private set; }
        public bool AttachAsPem { get; private set; }

        public CertExchangeDialog()
        {
            _certStore = new CertificateStore();
            _settings = ParclSettings.Load();
            InitializeComponents();
            LoadCertificates();
        }

        private void InitializeComponents()
        {
            Text = "Parcl - Send Certificate";
            Size = new Size(600, 450);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var headerLabel = new Label
            {
                Text = "Select the certificate you want to send to the recipient.\n" +
                       "This allows them to send you encrypted email.",
                Dock = DockStyle.Top,
                Height = 45,
                Padding = new Padding(8, 8, 8, 0)
            };

            _certListView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                GridLines = true
            };
            _certListView.Columns.Add("Subject", 200);
            _certListView.Columns.Add("Email", 150);
            _certListView.Columns.Add("Expires", 100);
            _certListView.Columns.Add("Key Usage", 100);
            _certListView.SelectedIndexChanged += CertList_SelectedChanged;

            _infoLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Padding = new Padding(8)
            };

            // Format options
            var formatPanel = new GroupBox
            {
                Text = "Attachment Format",
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(8)
            };

            _rbAttachCert = new RadioButton
            {
                Text = "DER (.cer) — Standard binary format",
                Location = new Point(12, 22),
                AutoSize = true,
                Checked = true
            };
            _rbAttachPem = new RadioButton
            {
                Text = "PEM (.pem) — Base64 text format",
                Location = new Point(300, 22),
                AutoSize = true
            };

            formatPanel.Controls.Add(_rbAttachCert);
            formatPanel.Controls.Add(_rbAttachPem);

            // Buttons
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 45,
                Padding = new Padding(8)
            };

            var cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
            var sendBtn = new Button { Text = "Attach to Message" };
            sendBtn.Click += SendButton_Click;

            buttonPanel.Controls.Add(cancelBtn);
            buttonPanel.Controls.Add(sendBtn);

            Controls.Add(_certListView);
            Controls.Add(headerLabel);
            Controls.Add(_infoLabel);
            Controls.Add(formatPanel);
            Controls.Add(buttonPanel);

            AcceptButton = sendBtn;
            CancelButton = cancelBtn;
        }

        private void LoadCertificates()
        {
            var certs = _certStore.GetEncryptionCertificates();
            foreach (var cert in certs)
            {
                var item = new ListViewItem(new[]
                {
                    cert.Subject,
                    cert.Email,
                    cert.NotAfter.ToString("yyyy-MM-dd"),
                    cert.IsEncryptionCert ? "Encryption" : "Signing"
                }) { Tag = cert };
                _certListView.Items.Add(item);
            }
        }

        private void CertList_SelectedChanged(object? sender, EventArgs e)
        {
            if (_certListView.SelectedItems.Count > 0)
            {
                var cert = (CertificateInfo)_certListView.SelectedItems[0].Tag;
                _infoLabel.Text = $"Issuer: {cert.Issuer} | Serial: {cert.SerialNumber} | " +
                                  $"Thumbprint: {cert.Thumbprint.Substring(0, 16)}...";
            }
        }

        private void SendButton_Click(object? sender, EventArgs e)
        {
            if (_certListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select a certificate to send.", "Parcl",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SelectedCertificate = (CertificateInfo)_certListView.SelectedItems[0].Tag;
            AttachAsPem = _rbAttachPem.Checked;
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
