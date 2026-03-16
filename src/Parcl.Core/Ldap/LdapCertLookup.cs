using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Parcl.Core.Config;
using Parcl.Core.Models;

namespace Parcl.Core.Ldap
{
    public class LdapCertLookup
    {
        private readonly ParclLogger? _logger;

        public LdapCertLookup() { }

        public LdapCertLookup(ParclLogger logger)
        {
            _logger = logger;
        }

        public Task<List<CertificateInfo>> LookupByEmailAsync(string email, LdapDirectoryEntry directory)
        {
            return Task.Run(() => LookupByEmail(email, directory));
        }

        /// <summary>
        /// Escapes special characters in an LDAP filter value per RFC 4515.
        /// </summary>
        public static string EscapeLdapFilter(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            var sb = new System.Text.StringBuilder(value.Length + 10);
            foreach (char c in value)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\5c"); break;
                    case '*':  sb.Append("\\2a"); break;
                    case '(':  sb.Append("\\28"); break;
                    case ')':  sb.Append("\\29"); break;
                    case '\0': sb.Append("\\00"); break;
                    default:   sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        public List<CertificateInfo> LookupByEmail(string email, LdapDirectoryEntry directory)
        {
            var results = new List<CertificateInfo>();
            var filter = string.Format(directory.SearchFilter, EscapeLdapFilter(email));
            var ldapPath = $"LDAP://{directory.Server}:{directory.Port}/{directory.BaseDn}";

            using (var entry = CreateDirectoryEntry(ldapPath, directory))
            using (var searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = filter;
                searcher.PropertiesToLoad.Add(directory.CertAttribute);
                searcher.PropertiesToLoad.Add("cn");
                searcher.PropertiesToLoad.Add("mail");

                var searchResults = searcher.FindAll();
                foreach (SearchResult result in searchResults)
                {
                    if (!result.Properties.Contains(directory.CertAttribute))
                        continue;

                    foreach (var certBytes in result.Properties[directory.CertAttribute])
                    {
                        if (certBytes is byte[] rawCert)
                        {
                            try
                            {
                                var cert = new X509Certificate2(rawCert);
                                var info = CertificateInfo.FromX509(cert);
                                if (info.IsValid)
                                    results.Add(info);
                            }
                            catch (Exception ex)
                            {
                                _logger?.Warn("LDAP", $"Skipping malformed certificate from {directory.Server}: {ex.Message}");
                            }
                        }
                    }
                }
            }

            return results;
        }

        public Task<List<CertificateInfo>> LookupAcrossDirectoriesAsync(
            string email, IEnumerable<LdapDirectoryEntry> directories)
        {
            return Task.Run(() =>
            {
                var allResults = new List<CertificateInfo>();
                foreach (var dir in directories)
                {
                    if (!dir.Enabled) continue;
                    try
                    {
                        var results = LookupByEmail(email, dir);
                        allResults.AddRange(results);
                    }
                    catch (Exception ex)
                    {
                        _logger?.Warn("LDAP", $"Directory lookup failed for {dir.Name} ({dir.Server}): {ex.Message}");
                    }
                }
                return allResults;
            });
        }

        public bool TestConnection(LdapDirectoryEntry directory)
        {
            try
            {
                var ldapPath = $"LDAP://{directory.Server}:{directory.Port}/{directory.BaseDn}";
                using (var entry = CreateDirectoryEntry(ldapPath, directory))
                {
                    // Force the connection by reading a property
                    _ = entry.NativeGuid;
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger?.Error("LDAP", $"Connection test failed for {directory.Server}:{directory.Port}: {ex.Message}");
                return false;
            }
        }

        private static DirectoryEntry CreateDirectoryEntry(string path, LdapDirectoryEntry config)
        {
            var entry = new DirectoryEntry(path);

            switch (config.AuthType)
            {
                case AuthType.Anonymous:
                    entry.AuthenticationType = AuthenticationTypes.Anonymous;
                    break;
                case AuthType.Simple:
                    entry.AuthenticationType = AuthenticationTypes.None;
                    if (config.BindDn != null)
                    {
                        entry.Username = config.BindDn;
                        entry.Password = config.GetBindPassword();
                    }
                    break;
                case AuthType.Negotiate:
                    entry.AuthenticationType = AuthenticationTypes.Secure;
                    break;
            }

            if (config.UseSsl)
                entry.AuthenticationType |= AuthenticationTypes.SecureSocketsLayer;

            return entry;
        }
    }
}
