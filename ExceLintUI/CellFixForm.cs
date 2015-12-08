using System;
using System.Windows.Forms;

namespace ExceLintUI
{
    public partial class CellFixForm : Form
    {
        Microsoft.Office.Interop.Excel.Range _cell;
        System.Drawing.Color _color;
        Action _fn;

        public CellFixForm(Microsoft.Office.Interop.Excel.Range cell, System.Drawing.Color color, Action ReAnalyzeFn)
        {
            _cell = cell;
            _color = color;
            _fn = ReAnalyzeFn;
            InitializeComponent();
        }

        private void FixButton_Click(object sender, EventArgs e)
        {
            // change the cell value
            _cell.Value2 = this.FixText.Text;

            // change color
            _cell.Interior.Color = _color;

            // close form
            this.Close();

            // call callback
            _fn.Invoke();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
