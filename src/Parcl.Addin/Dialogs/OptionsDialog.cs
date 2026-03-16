using System;
using System.Drawing;
using System.Windows.Forms;
using Parcl.Core.Config;
using Parcl.Core.Ldap;
using Parcl.Core.Models;

namespace Parcl.Addin.Dialogs
{
    public class OptionsDialog : Form
    {
        private readonly ParclSettings _settings;
        private TabControl _tabControl = null!;

        // LDAP controls
        private ListView _ldapListView = null!;
        private TextBox _ldapServer = null!;
        private TextBox _ldapPort = null!;
        private TextBox _ldapBaseDn = null!;
        private TextBox _ldapFilter = null!;
        private ComboBox _ldapAuth = null!;
        private TextBox _ldapBindDn = null!;
        private TextBox _ldapBindPassword = null!;
        private CheckBox _ldapSsl = null!;

        // Crypto controls
        private CheckBox _useNativeSmime = null!;
        private ComboBox _encAlgo = null!;
        private ComboBox _hashAlgo = null!;
        private ComboBox _certValidation = null!;
        private CheckBox _alwaysSign = null!;
        private CheckBox _alwaysEncrypt = null!;

        // Behavior controls
        private CheckBox _autoDecrypt = null!;
        private ComboBox _logLevel = null!;
        private ComboBox _autoLookup = null!;
        private CheckBox _promptMissing = null!;
        private CheckBox _showStatus = null!;

        // Cache controls
        private CheckBox _enableCache = null!;
        private NumericUpDown _cacheHours = null!;
        private NumericUpDown _maxCache = null!;

        public OptionsDialog()
        {
            _settings = ParclSettings.Load();
            InitializeComponents();
            LoadSettings();
        }

        private void InitializeComponents()
        {
            Text = "Parcl - Options";
            Size = new Size(600, 500);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            _tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(12, 4)
            };

            _tabControl.TabPages.Add(CreateBehaviorTab());
            _tabControl.TabPages.Add(CreateCryptoTab());
            _tabControl.TabPages.Add(CreateCacheTab());
            _tabControl.TabPages.Add(CreateLdapTab());

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

            Controls.Add(_tabControl);
            Controls.Add(buttonPanel);

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private TabPage CreateLdapTab()
        {
            var tab = new TabPage("LDAP (Optional)");
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 150
            };

            _ldapListView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true
            };
            _ldapListView.Columns.Add("Name", 120);
            _ldapListView.Columns.Add("Server", 150);
            _ldapListView.Columns.Add("Port", 60);
            _ldapListView.Columns.Add("SSL", 50);
            _ldapListView.SelectedIndexChanged += LdapList_SelectedChanged;

            var listButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 35
            };

            var addBtn = new Button { Text = "Add" };
            addBtn.Click += LdapAdd_Click;
            var removeBtn = new Button { Text = "Remove" };
            removeBtn.Click += LdapRemove_Click;
            var testBtn = new Button { Text = "Test Connection" };
            testBtn.Click += LdapTest_Click;

            listButtons.Controls.AddRange(new Control[] { addBtn, removeBtn, testBtn });

            splitContainer.Panel1.Controls.Add(_ldapListView);
            splitContainer.Panel1.Controls.Add(listButtons);

            // Detail panel
            var detailPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 8,
                Padding = new Padding(4)
            };

            _ldapServer = new TextBox { Dock = DockStyle.Fill };
            _ldapPort = new TextBox { Dock = DockStyle.Fill, Text = "636" };
            _ldapBaseDn = new TextBox { Dock = DockStyle.Fill };
            _ldapFilter = new TextBox { Dock = DockStyle.Fill, Text = "(mail={0})" };
            _ldapAuth = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _ldapAuth.Items.AddRange(new object[] { "Anonymous", "Simple", "Negotiate (Kerberos)" });
            _ldapBindDn = new TextBox { Dock = DockStyle.Fill };
            _ldapBindPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            _ldapSsl = new CheckBox { Text = "Use SSL/TLS", Checked = true };

            AddRow(detailPanel, 0, "Server:", _ldapServer);
            AddRow(detailPanel, 1, "Port:", _ldapPort);
            AddRow(detailPanel, 2, "Base DN:", _ldapBaseDn);
            AddRow(detailPanel, 3, "Search Filter:", _ldapFilter);
            AddRow(detailPanel, 4, "Authentication:", _ldapAuth);
            AddRow(detailPanel, 5, "Bind DN:", _ldapBindDn);
            AddRow(detailPanel, 6, "Bind Password:", _ldapBindPassword);
            detailPanel.Controls.Add(_ldapSsl, 1, 7);

            splitContainer.Panel2.Controls.Add(detailPanel);
            tab.Controls.Add(splitContainer);
            return tab;
        }

        private TabPage CreateCryptoTab()
        {
            var tab = new TabPage("Cryptography");
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 7,
                Height = 260,
                Padding = new Padding(12)
            };

            _useNativeSmime = new CheckBox
            {
                Text = "Use native Outlook S/MIME (compatible with Entrust, etc.)",
                AutoSize = true
            };

            var nativeHint = new Label
            {
                Text = "Native mode lets any S/MIME client decrypt. Parcl mode adds protected headers but requires Parcl.",
                ForeColor = System.Drawing.Color.Gray,
                Font = new System.Drawing.Font("Segoe UI", 8),
                AutoSize = true
            };

            _encAlgo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _encAlgo.Items.AddRange(new object[] { "AES-128-CBC", "AES-192-CBC", "AES-256-CBC" });

            _hashAlgo = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _hashAlgo.Items.AddRange(new object[] { "SHA-256", "SHA-384", "SHA-512" });

            _certValidation = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _certValidation.Items.AddRange(new object[] { "None", "Relaxed", "Strict" });

            _alwaysSign = new CheckBox { Text = "Always sign outgoing messages" };
            _alwaysEncrypt = new CheckBox { Text = "Always encrypt outgoing messages" };

            panel.Controls.Add(_useNativeSmime, 0, 0);
            panel.SetColumnSpan(_useNativeSmime, 2);
            panel.Controls.Add(nativeHint, 0, 1);
            panel.SetColumnSpan(nativeHint, 2);
            AddRow(panel, 2, "Encryption Algorithm:", _encAlgo);
            AddRow(panel, 3, "Hash Algorithm:", _hashAlgo);
            AddRow(panel, 4, "Certificate Validation:", _certValidation);
            panel.Controls.Add(_alwaysSign, 0, 5);
            panel.SetColumnSpan(_alwaysSign, 2);
            panel.Controls.Add(_alwaysEncrypt, 0, 6);
            panel.SetColumnSpan(_alwaysEncrypt, 2);

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateBehaviorTab()
        {
            var tab = new TabPage("General");
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 7,
                Height = 250,
                Padding = new Padding(12)
            };

            _autoDecrypt = new CheckBox { Text = "Automatically decrypt incoming messages" };
            _autoLookup = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _autoLookup.Items.AddRange(new object[] { "Manual", "On Compose", "On Send" });

            _promptMissing = new CheckBox { Text = "Prompt when recipient certificate not found" };
            _showStatus = new CheckBox { Text = "Show status bar in Outlook" };

            _logLevel = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _logLevel.Items.AddRange(new object[] { "Debug", "Info", "Warn", "Error" });

            var openLogsBtn = new Button { Text = "Open Log Folder", AutoSize = true };
            openLogsBtn.Click += (s, e) =>
            {
                var logDir = ParclAddIn.Current?.Logger?.GetLogDirectory()
                    ?? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Parcl", "logs");
                try { System.Diagnostics.Process.Start("explorer.exe", logDir); }
                catch { }
            };

            panel.Controls.Add(_autoDecrypt, 0, 0);
            panel.SetColumnSpan(_autoDecrypt, 2);
            AddRow(panel, 1, "Certificate lookup:", _autoLookup);
            panel.Controls.Add(_promptMissing, 0, 2);
            panel.SetColumnSpan(_promptMissing, 2);
            panel.Controls.Add(_showStatus, 0, 3);
            panel.SetColumnSpan(_showStatus, 2);
            AddRow(panel, 4, "Log level:", _logLevel);
            panel.Controls.Add(openLogsBtn, 1, 5);

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateCacheTab()
        {
            var tab = new TabPage("Cache");
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 4,
                Height = 160,
                Padding = new Padding(12)
            };

            _enableCache = new CheckBox { Text = "Enable certificate cache" };
            _cacheHours = new NumericUpDown { Minimum = 1, Maximum = 720, Value = 24, Dock = DockStyle.Fill };
            _maxCache = new NumericUpDown { Minimum = 10, Maximum = 10000, Value = 500, Dock = DockStyle.Fill };

            var clearBtn = new Button { Text = "Clear Cache" };
            clearBtn.Click += (s, e) =>
            {
                new CertificateCache().Clear();
                MessageBox.Show("Cache cleared.", "Parcl", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            panel.Controls.Add(_enableCache, 0, 0);
            panel.SetColumnSpan(_enableCache, 2);
            AddRow(panel, 1, "Expiration (hours):", _cacheHours);
            AddRow(panel, 2, "Max entries:", _maxCache);
            panel.Controls.Add(clearBtn, 1, 3);

            tab.Controls.Add(panel);
            return tab;
        }

        private void LoadSettings()
        {
            // Load LDAP directories
            foreach (var dir in _settings.LdapDirectories)
            {
                var item = new ListViewItem(new[]
                {
                    dir.Name, dir.Server, dir.Port.ToString(), dir.UseSsl ? "Yes" : "No"
                }) { Tag = dir };
                _ldapListView.Items.Add(item);
            }

            // Load crypto settings
            _useNativeSmime.Checked = _settings.Crypto.UseNativeSmime;
            _encAlgo.SelectedItem = _settings.Crypto.EncryptionAlgorithm;
            _hashAlgo.SelectedItem = _settings.Crypto.HashAlgorithm;
            _certValidation.SelectedIndex = (int)_settings.Crypto.ValidationMode;
            _alwaysSign.Checked = _settings.Crypto.AlwaysSign;
            _alwaysEncrypt.Checked = _settings.Crypto.AlwaysEncrypt;

            // Load behavior
            _autoDecrypt.Checked = _settings.Behavior.AutoDecrypt;
            _logLevel.SelectedItem = _settings.Behavior.LogLevel;
            if (_logLevel.SelectedIndex < 0) _logLevel.SelectedIndex = 1; // default Info
            _autoLookup.SelectedIndex = (int)_settings.Behavior.AutoLookup;
            _promptMissing.Checked = _settings.Behavior.PromptOnMissingCert;
            _showStatus.Checked = _settings.Behavior.ShowStatusBar;

            // Load cache
            _enableCache.Checked = _settings.Cache.EnableCertCache;
            _cacheHours.Value = _settings.Cache.CacheExpirationHours;
            _maxCache.Value = _settings.Cache.MaxCacheEntries;
        }

        private void OkButton_Click(object? sender, EventArgs e)
        {
            _settings.Crypto.UseNativeSmime = _useNativeSmime.Checked;
            _settings.Crypto.EncryptionAlgorithm = _encAlgo.SelectedItem?.ToString() ?? "AES-256-CBC";
            _settings.Crypto.HashAlgorithm = _hashAlgo.SelectedItem?.ToString() ?? "SHA-256";
            _settings.Crypto.ValidationMode = (CertValidationMode)_certValidation.SelectedIndex;
            _settings.Crypto.AlwaysSign = _alwaysSign.Checked;
            _settings.Crypto.AlwaysEncrypt = _alwaysEncrypt.Checked;

            _settings.Behavior.AutoDecrypt = _autoDecrypt.Checked;
            _settings.Behavior.LogLevel = _logLevel.SelectedItem?.ToString() ?? "Info";
            _settings.Behavior.AutoLookup = (LookupTrigger)_autoLookup.SelectedIndex;
            _settings.Behavior.PromptOnMissingCert = _promptMissing.Checked;
            _settings.Behavior.ShowStatusBar = _showStatus.Checked;

            _settings.Cache.EnableCertCache = _enableCache.Checked;
            _settings.Cache.CacheExpirationHours = (int)_cacheHours.Value;
            _settings.Cache.MaxCacheEntries = (int)_maxCache.Value;

            _settings.Save();

            // Apply log level change immediately
            if (ParclAddIn.Current?.Logger != null &&
                Enum.TryParse<Parcl.Core.Config.LogLevel>(
                    _settings.Behavior.LogLevel, true, out var newLevel))
            {
                ParclAddIn.Current.Logger.SetMinLevel(newLevel);
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void LdapList_SelectedChanged(object? sender, EventArgs e)
        {
            if (_ldapListView.SelectedItems.Count > 0)
            {
                var dir = (LdapDirectoryEntry)_ldapListView.SelectedItems[0].Tag;
                _ldapServer.Text = dir.Server;
                _ldapPort.Text = dir.Port.ToString();
                _ldapBaseDn.Text = dir.BaseDn;
                _ldapFilter.Text = dir.SearchFilter;
                _ldapAuth.SelectedIndex = (int)dir.AuthType;
                _ldapBindDn.Text = dir.BindDn ?? string.Empty;
                _ldapBindPassword.Text = dir.GetBindPassword();
                _ldapSsl.Checked = dir.UseSsl;
            }
        }

        private void LdapAdd_Click(object? sender, EventArgs e)
        {
            var dir = new LdapDirectoryEntry
            {
                Name = "New Directory",
                Server = _ldapServer.Text,
                Port = int.TryParse(_ldapPort.Text, out var p) ? p : 636,
                BaseDn = _ldapBaseDn.Text,
                SearchFilter = _ldapFilter.Text,
                AuthType = (AuthType)(_ldapAuth.SelectedIndex >= 0 ? _ldapAuth.SelectedIndex : 2),
                BindDn = string.IsNullOrWhiteSpace(_ldapBindDn.Text) ? null : _ldapBindDn.Text,
                BindPassword = null,
                UseSsl = _ldapSsl.Checked
            };

            if (!string.IsNullOrWhiteSpace(_ldapBindPassword.Text))
                dir.SetBindPassword(_ldapBindPassword.Text);

            _settings.LdapDirectories.Add(dir);
            var item = new ListViewItem(new[]
            {
                dir.Name, dir.Server, dir.Port.ToString(), dir.UseSsl ? "Yes" : "No"
            }) { Tag = dir };
            _ldapListView.Items.Add(item);
        }

        private void LdapRemove_Click(object? sender, EventArgs e)
        {
            if (_ldapListView.SelectedItems.Count > 0)
            {
                var dir = (LdapDirectoryEntry)_ldapListView.SelectedItems[0].Tag;
                _settings.LdapDirectories.Remove(dir);
                _ldapListView.SelectedItems[0].Remove();
            }
        }

        private void LdapTest_Click(object? sender, EventArgs e)
        {
            if (_ldapListView.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select a directory to test.", "Parcl");
                return;
            }

            var dir = (LdapDirectoryEntry)_ldapListView.SelectedItems[0].Tag;
            var lookup = new LdapCertLookup();
            var success = lookup.TestConnection(dir);

            MessageBox.Show(
                success ? "Connection successful!" : "Connection failed. Check server settings.",
                "Parcl - Connection Test",
                MessageBoxButtons.OK,
                success ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        private static void AddRow(TableLayoutPanel panel, int row, string label, Control control)
        {
            panel.Controls.Add(new Label
            {
                Text = label,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            }, 0, row);
            panel.Controls.Add(control, 1, row);
        }
    }
}
