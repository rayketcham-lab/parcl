using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace Parcl.Addin.TaskPane
{
    public class ParclTaskPaneHost : UserControl
    {
        private readonly ElementHost _host;
        private readonly ParclTaskPaneControl _wpfControl;

        public ParclTaskPaneHost()
        {
            _wpfControl = new ParclTaskPaneControl();
            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _wpfControl
            };
            Controls.Add(_host);
            Dock = DockStyle.Fill;
        }

        public ParclTaskPaneControl WpfControl => _wpfControl;
    }
}
