namespace Parcl.Core.Models
{
    public class LdapDirectoryEntry
    {
        public string Name { get; set; } = string.Empty;
        public string Server { get; set; } = string.Empty;
        public int Port { get; set; } = 389;
        public bool UseSsl { get; set; }
        public string BaseDn { get; set; } = string.Empty;
        public string SearchFilter { get; set; } = "(mail={0})";
        public string CertAttribute { get; set; } = "userCertificate;binary";
        public AuthType AuthType { get; set; } = AuthType.Negotiate;
        public string? BindDn { get; set; }
        public bool Enabled { get; set; } = true;

        public string ConnectionString =>
            $"{(UseSsl ? "ldaps" : "ldap")}://{Server}:{Port}";
    }

    public enum AuthType
    {
        Anonymous,
        Simple,
        Negotiate
    }
}
