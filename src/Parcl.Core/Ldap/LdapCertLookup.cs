using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Parcl.Core.Models;

namespace Parcl.Core.Ldap
{
    public class LdapCertLookup
    {
        public Task<List<CertificateInfo>> LookupByEmailAsync(string email, LdapDirectoryEntry directory)
        {
            return Task.Run(() => LookupByEmail(email, directory));
        }

        public List<CertificateInfo> LookupByEmail(string email, LdapDirectoryEntry directory)
        {
            var results = new List<CertificateInfo>();
            var filter = string.Format(directory.SearchFilter, email);
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
                            catch (Exception)
                            {
                                // Skip malformed certificates
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
                    catch (Exception)
                    {
                        // Log and continue to next directory
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
            catch
            {
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
                        entry.Username = config.BindDn;
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
