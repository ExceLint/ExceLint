using System;
using System.Windows.Forms;

namespace ExceLintUI
{
    /// <summary>
    /// This progress bar's lifecycle should be managed by the UI layer.
    /// </summary>
    public partial class ProgBar : Form
    {
        private int _count = 0;
        private Boolean _cancel = false;

        public ProgBar()
        {
            InitializeComponent();

            workProgress.Minimum = 0;
            workProgress.Maximum = 100;
            this.Visible = true;
        }

        public void IncrementProgress()
        {
            // if this method is called from any thread other than
            // the GUI thread, call the method on the correct thread
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(this.IncrementProgress));
                return;
            }

            int pbval;

            if (_count < 0)
            {
                pbval = 0;
            }
            else if (_count > workProgress.Maximum)
            {
                pbval = workProgress.Maximum;
            }
            else
            {
                pbval = (int)(_count);
            }

            workProgress.Value = pbval;

            _count++;
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {

        }

        private void ProgBar_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _cancel = true;
        }

        public bool IsCancelled()
        {
            return _cancel;
        }
    }
}
