using System;
using System.Windows.Forms;

namespace ExceLintUI
{
    public partial class OverwriteOrAppend : Form
    {
        public OverwriteOrAppend()
        {
            StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        public void SetText(string text)
        {
            messageText.Text = text;
        }

        public string Message
        {
            get { return messageText.Text; }
            set { messageText.Text = value; }
        }

        private void appendButton_Click_1(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void overwriteButton_Click_1(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
