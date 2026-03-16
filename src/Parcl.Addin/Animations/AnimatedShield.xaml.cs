using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace Parcl.Addin.Animations
{
    public partial class AnimatedShield : UserControl
    {
        public AnimatedShield()
        {
            InitializeComponent();
        }

        public void PlaySign()
        {
            if (TryFindResource("SignAnimation") is Storyboard sb)
                sb.Begin(this);
        }

        public void PlayVerifyFail()
        {
            if (TryFindResource("VerifyFailAnimation") is Storyboard sb)
                sb.Begin(this);
        }

        public void StartIdlePulse()
        {
            if (TryFindResource("IdlePulse") is Storyboard sb)
                sb.Begin(this);
        }
    }
}
