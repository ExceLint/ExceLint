using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExceLintCLIGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // read defaults
            string iDir = (string)Properties.Settings.Default["excelintrunnerPathDefault"];
        }

        private void spectralCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (spectralCheckBox.Checked)
            {
                sheetsCheckbox.Checked = true;
                allcellsCheckBox.Checked = false;
                columnsCheckBox.Checked = false;
                rowsCheckBox.Checked = false;
                levelsCheckBox.Checked = false;

                sheetsCheckbox.Enabled = false;
                allcellsCheckBox.Enabled = false;
                columnsCheckBox.Enabled = false;
                rowsCheckBox.Enabled = false;
                levelsCheckBox.Enabled = false;
            } else
            {
                sheetsCheckbox.Enabled = true;
                allcellsCheckBox.Enabled = true;
                columnsCheckBox.Enabled = true;
                rowsCheckBox.Enabled = true;
                levelsCheckBox.Enabled = true;
            }
        }
    }
}
