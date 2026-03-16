using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace Parcl.Addin.Animations
{
    public partial class AnimatedLock : UserControl
    {
        private bool _isLocked;

        public AnimatedLock()
        {
            InitializeComponent();
        }

        public bool IsLocked
        {
            get => _isLocked;
            set
            {
                if (_isLocked == value) return;
                _isLocked = value;
                PlayAnimation(value ? "LockAnimation" : "UnlockAnimation");
            }
        }

        public void PlayLock() => IsLocked = true;
        public void PlayUnlock() => IsLocked = false;

        private void PlayAnimation(string name)
        {
            if (TryFindResource(name) is Storyboard sb)
                sb.Begin(this);
        }
    }
}
