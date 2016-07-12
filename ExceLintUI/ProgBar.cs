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
        private Action _cancel_action;

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

        public void registerCancelCallback(Action cancelAction)
        {
            _cancel_action = cancelAction;
        }

        private void ProgBar_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MessageBox.Show("You clicked me!");
            if (_cancel_action != null)
            {
                _cancel_action();
            }
        }
    }
}
