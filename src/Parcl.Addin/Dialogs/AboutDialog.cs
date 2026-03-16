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
            Size = new Size(420, 380);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(30, 30, 35);

            // Logo / Title
            var titleLabel = new Label
            {
                Text = "PARCL",
                Font = new Font("Segoe UI", 28, FontStyle.Bold),
                ForeColor = Color.FromArgb(79, 195, 247),
                AutoSize = true,
                Location = new Point(24, 16)
            };

            var subtitleLabel = new Label
            {
                Text = "S/MIME Certificate Manager for Outlook",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(180, 180, 190),
                AutoSize = true,
                Location = new Point(26, 58)
            };

            var versionLabel = new Label
            {
                Text = $"Version {version?.Major ?? 1}.{version?.Minor ?? 2}.{version?.Build ?? 0}",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(130, 130, 140),
                AutoSize = true,
                Location = new Point(26, 80)
            };

            // Separator
            var sep = new Label
            {
                BorderStyle = BorderStyle.Fixed3D,
                Height = 2,
                Width = 370,
                Location = new Point(24, 106)
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
                Location = new Point(26, 118)
            };

            // Links
            var ghLink = CreateLink("GitHub: rayketcham-lab/parcl",
                "https://github.com/rayketcham-lab/parcl", new Point(26, 178));

            var webLink = CreateLink("quantumnexum.com",
                "https://quantumnexum.com", new Point(26, 200));

            var supportLink = CreateLink("Support: help@quantumnexum.com",
                "mailto:help@quantumnexum.com", new Point(26, 222));

            // License
            var licenseLabel = new Label
            {
                Text = "Licensed under the MIT License",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 110),
                AutoSize = true,
                Location = new Point(26, 256)
            };

            var copyrightLabel = new Label
            {
                Text = $"\u00A9 {DateTime.Now.Year} Quantum Nexum. All rights reserved.",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 110),
                AutoSize = true,
                Location = new Point(26, 274)
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
                Location = new Point(310, 300)
            };
            closeBtn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 80);

            Controls.AddRange(new Control[]
            {
                titleLabel, subtitleLabel, versionLabel, sep,
                infoLabel, ghLink, webLink, supportLink,
                licenseLabel, copyrightLabel, closeBtn
            });

            AcceptButton = closeBtn;
            CancelButton = closeBtn;
        }

        private static LinkLabel CreateLink(string text, string url, Point location)
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
                try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
                catch { }
            };
            return link;
        }
    }
}
