using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace Parcl.Addin.Animations
{
    public enum CertStatus { Valid, Expired, Warning }

    public partial class CertificateBadge : UserControl
    {
        public CertificateBadge()
        {
            InitializeComponent();
        }

        public void SetStatus(CertStatus status)
        {
            var animName = status switch
            {
                CertStatus.Valid => "ValidAnimation",
                CertStatus.Expired => "ExpiredAnimation",
                CertStatus.Warning => "WarningAnimation",
                _ => "ValidAnimation"
            };

            if (TryFindResource(animName) is Storyboard sb)
                sb.Begin(this);
        }
    }
}
