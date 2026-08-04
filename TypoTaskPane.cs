using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace AdrkhaTypograph
{
    public partial class TypoTaskPane : UserControl
    {
        private ElementHost _elementHost;
        private TypoControl _typoControl;

        public TypoTaskPane()
        {
            InitializeComponent();
            InitializeWpfControl();
        }

        private void InitializeWpfControl()
        {
            try
            {
                _elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill
                };

                _typoControl = new TypoControl();
                _elementHost.Child = _typoControl;

                this.Controls.Add(_elementHost);
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطأ أثناء تحميل واجهة WPF:\n" + ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}