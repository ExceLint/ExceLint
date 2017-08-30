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

            this.CenterToParent();

            workProgress.Minimum = 0;
            workProgress.Maximum = 100;
            this.Visible = true;
        }

        public void IncrementProgress()
        {
            IncrementProgressN(1);
        }

        public void IncrementProgressN(int n)
        {
            // if this method is called from any thread other than
            // the GUI thread, call the method on the correct thread
            if (this.InvokeRequired)
            {
                BeginInvoke(new Action<int>(IncrementProgressN), n);
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

            _count += n;
        }

        public void Reset()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(this.Reset));
                return;
            }

            _count = 0;
            workProgress.Value = 0;
        }

        public void registerCancelCallback(Action cancelAction)
        {
            _cancel_action = cancelAction;
        }

        public void GoAway()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(this.GoAway));
                return;
            }

            this.Dispose();
        }

        private void ProgBar_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _cancel_action?.Invoke();
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {
            this.FormClosing += new FormClosingEventHandler(ProgBar_Closing);
        }
    }
}
