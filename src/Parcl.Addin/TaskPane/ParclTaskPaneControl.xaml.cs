using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Parcl.Addin.Animations;
using Parcl.Addin.Dialogs;
using Parcl.Core.Config;
using Parcl.Core.Crypto;
using Parcl.Core.Ldap;
using Parcl.Core.Models;

namespace Parcl.Addin.TaskPane
{
    public partial class ParclTaskPaneControl : UserControl
    {
        private readonly ParclSettings _settings;
        private readonly CertificateStore _certStore;
        private readonly LdapCertLookup _ldapLookup;
        private readonly CertificateCache _certCache;
        private readonly ParclLogger _logger;

        public ParclTaskPaneControl()
        {
            InitializeComponent();
            _settings = ParclSettings.Load();
            _certStore = new CertificateStore();
            _ldapLookup = new LdapCertLookup();
            _certCache = new CertificateCache(
                _settings.Cache.CacheExpirationHours,
                _settings.Cache.MaxCacheEntries);
            _logger = new ParclLogger();

            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _logger.Info("TaskPane", "Parcl task pane loaded");
            LogoShield.StartIdlePulse();
            RefreshCertificateDisplay();
            UpdateStatus("Parcl ready — certificates loaded");
        }

        private void RefreshCertificateDisplay()
        {
            try
            {
                if (!string.IsNullOrEmpty(_settings.UserProfile.SigningCertThumbprint))
                {
                    var cert = _certStore.FindByThumbprint(_settings.UserProfile.SigningCertThumbprint);
                    if (cert != null)
                    {
                        var info = CertificateInfo.FromX509(cert);
                        SignCertLabel.Text = $"Signing: {info.Subject}";
                        SigningBadge.SetStatus(info.IsValid ? CertStatus.Valid :
                            info.IsExpired ? CertStatus.Expired : CertStatus.Warning);
                        SigningStatus.Text = info.IsValid ? "Valid" : "Expired";
                        _logger.Info("Certs", $"Signing cert loaded: {info.Thumbprint.Substring(0, 8)}");
                    }
                }

                if (!string.IsNullOrEmpty(_settings.UserProfile.EncryptionCertThumbprint))
                {
                    var cert = _certStore.FindByThumbprint(_settings.UserProfile.EncryptionCertThumbprint);
                    if (cert != null)
                    {
                        var info = CertificateInfo.FromX509(cert);
                        EncCertLabel.Text = $"Encryption: {info.Subject}";
                        EncryptionBadge.SetStatus(info.IsValid ? CertStatus.Valid :
                            info.IsExpired ? CertStatus.Expired : CertStatus.Warning);
                        EncryptionStatus.Text = info.IsValid ? "Valid" : "Expired";
                        _logger.Info("Certs", $"Encryption cert loaded: {info.Thumbprint.Substring(0, 8)}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Certs", $"Failed to load certificates: {ex.Message}");
                UpdateStatus("Error loading certificates");
            }
        }

        private void BtnSelectCerts_Click(object sender, RoutedEventArgs e)
        {
            _logger.Debug("UI", "Certificate selector opened");
            using (var dialog = new CertificateSelectorDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _settings.Reload();
                    RefreshCertificateDisplay();
                    UpdateStatus("Certificates updated");
                    _logger.Info("Certs", "Certificate selection updated by user");
                }
            }
        }

        private void BtnEncrypt_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("Crypto", "Encrypt action triggered from task pane");
            EncryptionLock.PlayLock();
            AnimateButtonFlash(BtnEncrypt, "#00E676");
            UpdateStatus("Encrypting message...");
        }

        private void BtnDecrypt_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("Crypto", "Decrypt action triggered from task pane");
            EncryptionLock.PlayUnlock();
            AnimateButtonFlash(BtnDecrypt, "#4FC3F7");
            UpdateStatus("Decrypting message...");
        }

        private void BtnSign_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("Crypto", "Sign action triggered from task pane");
            SigningShield.PlaySign();
            AnimateButtonFlash(BtnSign, "#00E676");
            UpdateStatus("Signing message...");
        }

        private void BtnExchange_Click(object sender, RoutedEventArgs e)
        {
            _logger.Info("Exchange", "Certificate exchange triggered from task pane");
            using (var dialog = new CertExchangeDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    AnimateButtonFlash(BtnExchange, "#FFD740");
                    UpdateStatus($"Certificate attached for exchange");
                    _logger.Info("Exchange", $"Certificate selected for exchange: {dialog.SelectedCertificate?.Thumbprint?.Substring(0, 8)}");
                }
            }
        }

        private async void BtnLookup_Click(object sender, RoutedEventArgs e)
        {
            var email = LookupEmailBox.Text.Trim();
            if (string.IsNullOrEmpty(email) || email == "user@example.com")
            {
                UpdateStatus("Enter an email address to search");
                return;
            }

            _logger.Info("LDAP", $"Certificate lookup initiated for: {email}");

            LookupSpinnerPanel.Visibility = Visibility.Visible;
            LookupResults.Visibility = Visibility.Collapsed;
            LookupResults.Items.Clear();
            UpdateStatus($"Searching for {email}...");

            try
            {
                // Check cache first
                var cached = _certCache.Get(email);
                if (cached != null)
                {
                    _logger.Debug("LDAP", $"Cache hit for {email}: {cached.Count} cert(s)");
                    DisplayLookupResults(cached, email, fromCache: true);
                    return;
                }

                _logger.Debug("LDAP", $"Cache miss for {email}, querying {_settings.LdapDirectories.Count} director(ies)");

                var results = await _ldapLookup.LookupAcrossDirectoriesAsync(
                    email, _settings.LdapDirectories);

                if (results.Count > 0)
                    _certCache.Add(email, results);

                DisplayLookupResults(results, email, fromCache: false);
            }
            catch (Exception ex)
            {
                _logger.Error("LDAP", $"Lookup failed for {email}: {ex.Message}");
                LookupSpinnerPanel.Visibility = Visibility.Collapsed;
                UpdateStatus($"Lookup failed: {ex.Message}");
            }
        }

        private void DisplayLookupResults(List<CertificateInfo> results, string email, bool fromCache)
        {
            LookupSpinnerPanel.Visibility = Visibility.Collapsed;
            LookupResults.Visibility = Visibility.Visible;

            if (results.Count == 0)
            {
                LookupResults.Items.Add("No certificates found");
                UpdateStatus($"No certificates found for {email}");
                _logger.Info("LDAP", $"No results for {email}");
            }
            else
            {
                foreach (var cert in results)
                {
                    LookupResults.Items.Add(
                        $"{cert.Subject} | Expires: {cert.NotAfter:yyyy-MM-dd} | {cert.Thumbprint.Substring(0, 8)}...");
                }
                var cacheNote = fromCache ? " (cached)" : "";
                UpdateStatus($"Found {results.Count} certificate(s) for {email}{cacheNote}");
                _logger.Info("LDAP", $"Found {results.Count} cert(s) for {email}{cacheNote}");
            }
        }

        private void LookupEmailBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (LookupEmailBox.Text == "user@example.com")
                LookupEmailBox.Text = "";
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = message;
            if (TryFindResource("StatusFade") is Storyboard sb)
                sb.Begin(this);
            _logger.Debug("UI", $"Status: {message}");
        }

        private void AnimateButtonFlash(Button button, string hexColor)
        {
            var color = (Color)ColorConverter.ConvertFromString(hexColor);
            var animation = new ColorAnimation
            {
                To = color,
                Duration = TimeSpan.FromSeconds(0.3),
                AutoReverse = true
            };

            var brush = new SolidColorBrush(Colors.Transparent);
            button.Background = brush;
            brush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }
    }
}
