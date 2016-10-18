using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExceLintUI
{
    public partial class OverwriteOrAppend : Form
    {
        public OverwriteOrAppend()
        {
            InitializeComponent();
        }

        public void SetText(string text)
        {
            messageText.Text = text;
        }
    }
}
