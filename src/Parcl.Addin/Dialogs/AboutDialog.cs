using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace Parcl.Addin.Dialogs
{
    public class AboutDialog : Form
    {
        public AboutDialog()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;

            Text = "About Parcl";
            Size = new Size(460, 490);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(30, 30, 35);

            // Logo / Title — sized to not overlap subtitle
            var titleLabel = new Label
            {
                Text = "PARCL",
                Font = new Font("Segoe UI", 22, FontStyle.Bold),
                ForeColor = Color.FromArgb(79, 195, 247),
                AutoSize = true,
                Location = new Point(24, 10)
            };

            var subtitleLabel = new Label
            {
                Text = "S/MIME Certificate Manager for Outlook",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(180, 180, 190),
                AutoSize = true,
                Location = new Point(26, 52)
            };

            bool fipsEnabled = false;
            try { fipsEnabled = System.Security.Cryptography.CryptoConfig.AllowOnlyFipsAlgorithms; }
            catch { }

            var versionLabel = new Label
            {
                Text = $"Version {System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()?.InformationalVersion?.Split('+')[0] ?? $"{version?.Major}.{version?.Minor}.{version?.Build}"}" +
                       (fipsEnabled ? "  |  FIPS Mode" : ""),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(130, 130, 140),
                AutoSize = true,
                Location = new Point(26, 74)
            };

            // Separator
            var sep = new Label
            {
                BorderStyle = BorderStyle.Fixed3D,
                Height = 2,
                Width = 410,
                Location = new Point(24, 98)
            };

            // Info section
            var infoLabel = new Label
            {
                Text = "Encrypt, sign, and manage S/MIME certificates\n" +
                       "directly from Outlook. Supports LDAP lookup,\n" +
                       "certificate exchange, and at-rest encryption.",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(200, 200, 210),
                AutoSize = true,
                Location = new Point(26, 110)
            };

            // Links
            var ghLink = CreateLink("GitHub: rayketcham-lab/parcl",
                "https://github.com/rayketcham-lab/parcl", new Point(26, 170));

            var webLink = CreateLink("quantumnexum.com",
                "https://quantumnexum.com", new Point(26, 192));

            var supportLink = CreateLink("Support: help@quantumnexum.com",
                "mailto:help@quantumnexum.com", new Point(26, 214));

            // GitHub action buttons
            var reportBtn = new Button
            {
                Text = "Report Issue",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 50, 60),
                ForeColor = Color.FromArgb(79, 195, 247),
                Font = new Font("Segoe UI", 8),
                Size = new Size(100, 26),
                Location = new Point(26, 244)
            };
            reportBtn.FlatAppearance.BorderColor = Color.FromArgb(79, 195, 247);
            reportBtn.Click += (s, e) =>
            {
                var sysInfo = $"Parcl v{version?.Major}.{version?.Minor}.{version?.Build}";
                var url = "https://github.com/rayketcham-lab/parcl/issues/new"
                    + $"?title=Bug:+&body=%23%23+Description%0A%0A%23%23+Steps+to+Reproduce%0A%0A%23%23+System%0A{Uri.EscapeDataString(sysInfo)}";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch { }
            };

            var suggestBtn = new Button
            {
                Text = "Suggest Feature",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 50, 60),
                ForeColor = Color.FromArgb(129, 199, 132),
                Font = new Font("Segoe UI", 8),
                Size = new Size(110, 26),
                Location = new Point(134, 244)
            };
            suggestBtn.FlatAppearance.BorderColor = Color.FromArgb(129, 199, 132);
            suggestBtn.Click += (s, e) =>
            {
                var url = "https://github.com/rayketcham-lab/parcl/issues/new"
                    + "?title=feat:+&body=%23%23+Feature+Request%0A%0A%23%23+Use+Case%0A%0A%23%23+Expected+Behavior";
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch { }
            };

            var viewIssuesLink = CreateLink("View all issues on GitHub",
                "https://github.com/rayketcham-lab/parcl/issues", new Point(26, 278));

            // License
            var licenseLabel = new Label
            {
                Text = "Licensed under the Apache License 2.0",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 110),
                AutoSize = true,
                Location = new Point(26, 310)
            };

            var copyrightLabel = new Label
            {
                Text = $"\u00A9 {DateTime.Now.Year} Quantum Nexum. All rights reserved.",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 110),
                AutoSize = true,
                Location = new Point(26, 328)
            };

            // Close button
            var closeBtn = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(50, 50, 60),
                ForeColor = Color.FromArgb(200, 200, 210),
                Font = new Font("Segoe UI", 9),
                Size = new Size(80, 30),
                Location = new Point(350, 370)
            };
            closeBtn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 80);

            Controls.AddRange(new Control[]
            {
                titleLabel, subtitleLabel, versionLabel, sep,
                infoLabel, ghLink, webLink, supportLink,
                reportBtn, suggestBtn, viewIssuesLink,
                licenseLabel, copyrightLabel, closeBtn
            });

            AcceptButton = closeBtn;
            CancelButton = closeBtn;
        }

        private LinkLabel CreateLink(string text, string url, Point location)
        {
            var link = new LinkLabel
            {
                Text = text,
                Font = new Font("Segoe UI", 9),
                LinkColor = Color.FromArgb(79, 195, 247),
                ActiveLinkColor = Color.FromArgb(129, 212, 250),
                VisitedLinkColor = Color.FromArgb(79, 195, 247),
                AutoSize = true,
                Location = location
            };
            link.LinkClicked += (s, e) =>
            {
                // mailto: links reenter Outlook which deadlocks against this modal dialog.
                // Close the dialog first, then launch the link after a short delay.
                if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    var timer = new Timer { Interval = 150 };
                    timer.Tick += (ts, te) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                        catch { }
                    };
                    timer.Start();
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                    catch { }
                }
            };
            return link;
        }
    }
}
