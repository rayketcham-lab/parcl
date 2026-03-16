using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using Parcl.Core.Crypto;
using Parcl.Core.Models;

namespace Parcl.Addin.Dialogs
{
    public class CertContactsDialog : Form
    {
        private ListView _listView = null!;
        private Button _importBtn = null!;
        private Button _removeBtn = null!;
        private Button _refreshBtn = null!;
        private Button _closeBtn = null!;

        public CertContactsDialog()
        {
            InitializeComponents();
            LoadCertificates();
        }

        private void InitializeComponents()
        {
            Text = "Parcl - Certificate Contacts";
            Size = new Size(700, 500);
            MinimumSize = new Size(500, 350);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(30, 30, 35);

            // ListView
            _listView = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                GridLines = false,
                BackColor = Color.FromArgb(40, 40, 48),
                ForeColor = Color.FromArgb(200, 200, 210),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Location = new Point(12, 12),
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };

            _listView.Columns.Add("Email", 180);
            _listView.Columns.Add("Subject", 200);
            _listView.Columns.Add("Status", 90);
            _listView.Columns.Add("Expires", 90);
            _listView.Columns.Add("Thumbprint", 100);

            _listView.OwnerDraw = true;
            _listView.DrawColumnHeader += ListView_DrawColumnHeader;
            _listView.DrawItem += ListView_DrawItem;
            _listView.DrawSubItem += ListView_DrawSubItem;
            _listView.SelectedIndexChanged += ListView_SelectedIndexChanged;

            // Button panel
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.FromArgb(30, 30, 35),
                Padding = new Padding(12, 8, 12, 8)
            };

            _importBtn = CreateButton("Import...", 0);
            _importBtn.Click += OnImportClick;

            _removeBtn = CreateButton("Remove", 1);
            _removeBtn.Enabled = false;
            _removeBtn.Click += OnRemoveClick;

            _refreshBtn = CreateButton("Refresh", 2);
            _refreshBtn.Click += OnRefreshClick;

            _closeBtn = CreateButton("Close", -1);
            _closeBtn.DialogResult = DialogResult.OK;
            _closeBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            buttonPanel.Controls.AddRange(new Control[] { _importBtn, _removeBtn, _refreshBtn, _closeBtn });

            Controls.Add(_listView);
            Controls.Add(buttonPanel);

            AcceptButton = _closeBtn;
            CancelButton = _closeBtn;

            Resize += OnFormResize;
            UpdateListViewSize();
        }

        private Button CreateButton(string text, int positionIndex)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 50, 60),
                ForeColor = Color.FromArgb(200, 200, 210),
                Font = new Font("Segoe UI", 9),
                Size = new Size(90, 30)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 80);

            if (positionIndex >= 0)
            {
                btn.Location = new Point(12 + positionIndex * 100, 10);
                btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            }
            else
            {
                // Right-aligned (Close button)
                btn.Location = new Point(ClientSize.Width - 102, 10);
                btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            }

            return btn;
        }

        private void UpdateListViewSize()
        {
            _listView.Size = new Size(ClientSize.Width - 24, ClientSize.Height - 74);
        }

        private void OnFormResize(object sender, EventArgs e)
        {
            UpdateListViewSize();
        }

        private void ListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            _removeBtn.Enabled = _listView.SelectedItems.Count > 0;
        }

        private void ListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            using (var brush = new SolidBrush(Color.FromArgb(45, 45, 55)))
                e.Graphics.FillRectangle(brush, e.Bounds);
            using (var brush = new SolidBrush(Color.FromArgb(160, 160, 170)))
            using (var font = new Font("Segoe UI", 8.5f, FontStyle.Bold))
            {
                var rect = new Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height);
                var sf = new StringFormat { LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString(e.Header.Text, font, brush, rect, sf);
            }
        }

        private void ListView_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            // Handled in DrawSubItem
        }

        private void ListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            Color bgColor = e.Item.Selected
                ? Color.FromArgb(55, 55, 70)
                : (e.ItemIndex % 2 == 0
                    ? Color.FromArgb(40, 40, 48)
                    : Color.FromArgb(36, 36, 43));

            using (var brush = new SolidBrush(bgColor))
                e.Graphics.FillRectangle(brush, e.Bounds);

            Color textColor;
            if (e.ColumnIndex == 2) // Status column
            {
                string status = e.SubItem.Text;
                if (status == "Expired")
                    textColor = Color.FromArgb(239, 83, 80);
                else if (status == "Expiring Soon")
                    textColor = Color.FromArgb(255, 167, 38);
                else
                    textColor = Color.FromArgb(102, 187, 106);
            }
            else
            {
                textColor = Color.FromArgb(200, 200, 210);
            }

            using (var brush = new SolidBrush(textColor))
            using (var font = new Font("Segoe UI", 9))
            {
                var rect = new Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height);
                var sf = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap
                };
                e.Graphics.DrawString(e.SubItem.Text, font, brush, rect, sf);
            }
        }

        private void LoadCertificates()
        {
            _listView.Items.Clear();

            using (var store = new X509Store(StoreName.AddressBook, StoreLocation.CurrentUser))
            {
                try
                {
                    store.Open(OpenFlags.ReadOnly);

                    foreach (var cert in store.Certificates.Cast<X509Certificate2>())
                    {
                        var info = CertificateInfo.FromX509(cert);

                        string status;
                        if (info.IsExpired)
                            status = "Expired";
                        else if ((info.NotAfter - DateTime.UtcNow).TotalDays < 30)
                            status = "Expiring Soon";
                        else
                            status = "Valid";

                        string email = !string.IsNullOrEmpty(info.Email)
                            ? info.Email
                            : "(none)";

                        string thumbShort = info.Thumbprint.Length >= 8
                            ? info.Thumbprint.Substring(0, 8)
                            : info.Thumbprint;

                        var item = new ListViewItem(email);
                        item.SubItems.Add(info.Subject);
                        item.SubItems.Add(status);
                        item.SubItems.Add(info.NotAfter.ToString("yyyy-MM-dd"));
                        item.SubItems.Add(thumbShort);
                        item.Tag = info.Thumbprint;

                        _listView.Items.Add(item);
                    }
                }
                finally
                {
                    store.Close();
                }
            }
        }

        private void OnImportClick(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = "Import Certificate";
                ofd.Filter = "Certificate Files|*.cer;*.pem;*.crt;*.der|All Files|*.*";

                if (ofd.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    byte[] rawData = File.ReadAllBytes(ofd.FileName);

                    // Handle PEM format: strip headers and base64 decode
                    string text = System.Text.Encoding.ASCII.GetString(rawData);
                    if (text.Contains("-----BEGIN CERTIFICATE-----"))
                    {
                        string base64 = text
                            .Replace("-----BEGIN CERTIFICATE-----", "")
                            .Replace("-----END CERTIFICATE-----", "")
                            .Trim();
                        rawData = Convert.FromBase64String(base64);
                    }

                    var cert = new X509Certificate2(rawData);

                    using (var store = new X509Store(StoreName.AddressBook, StoreLocation.CurrentUser))
                    {
                        store.Open(OpenFlags.ReadWrite);
                        store.Add(cert);
                        store.Close();
                    }

                    LoadCertificates();

                    MessageBox.Show(
                        $"Certificate imported successfully.\n\nSubject: {cert.Subject}",
                        "Parcl", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Failed to import certificate:\n{ex.Message}",
                        "Parcl", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void OnRemoveClick(object sender, EventArgs e)
        {
            if (_listView.SelectedItems.Count == 0)
                return;

            var selected = _listView.SelectedItems[0];
            string thumbprint = (string)selected.Tag;

            var result = MessageBox.Show(
                $"Remove certificate for {selected.Text}?",
                "Parcl - Confirm Remove",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                using (var store = new X509Store(StoreName.AddressBook, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadWrite);
                    var found = store.Certificates.Find(
                        X509FindType.FindByThumbprint, thumbprint, false);
                    if (found.Count > 0)
                        store.Remove(found[0]);
                    store.Close();
                }

                LoadCertificates();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to remove certificate:\n{ex.Message}",
                    "Parcl", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnRefreshClick(object sender, EventArgs e)
        {
            LoadCertificates();
        }
    }
}
