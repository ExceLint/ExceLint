using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExceLintUI
{
    /// <summary>
    /// This progress bar's lifecycle should be managed by the UI layer.
    /// </summary>
    public partial class ProgBar : Form
    {
        private int _count = 0;

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
    }
}
