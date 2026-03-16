namespace Parcl.Core.Models
{
    public class UserProfile
    {
        public string EmailAddress { get; set; } = string.Empty;
        public string? SigningCertThumbprint { get; set; }
        public string? EncryptionCertThumbprint { get; set; }
        public string DisplayName { get; set; } = string.Empty;
    }
}
