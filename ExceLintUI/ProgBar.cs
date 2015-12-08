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
    public class ProgressMaxUnsetException : Exception { }

    /// <summary>
    /// This progress bar's lifecycle should be managed by the UI layer.
    /// </summary>
    public partial class ProgBar : Form
    {
        private bool _max_set = false;
        private int _count = 0;

        public ProgBar()
        {
            InitializeComponent();
            workProgress.Minimum = 0;
            workProgress.Maximum = 100;
            this.Visible = true;
        }

        private void ProgBar_Load(object sender, EventArgs e)
        {

        }

        public void IncrementProgress()
        {
            // if this method is called from any thread other than
            // the GUI thread, call the method on the correct thread
            if (workProgress.InvokeRequired)
            {
                workProgress.Invoke(new MethodInvoker(() => IncrementProgress()));
                return;
            }

            if (!_max_set)
            {
                throw new ProgressMaxUnsetException();
            }

            if (_count < 0)
            {
                workProgress.Value = 0;
            }
            else if (_count > workProgress.Maximum)
            {
                workProgress.Value = workProgress.Maximum;
            }
            else
            {
                workProgress.Value = (int)(_count);
            }
            _count++;
        }

        public int maxProgress()
        {
            if (!_max_set)
            {
                throw new ProgressMaxUnsetException();
            }
            return workProgress.Maximum;
        }

        public void setMax(int max_updates)
        {
            workProgress.Maximum = max_updates;
            _max_set = true;
        }
    }
}
