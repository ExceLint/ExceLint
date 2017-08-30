namespace ExceLintUI
{
    partial class ProgBar
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.workProgress = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // workProgress
            // 
            this.workProgress.Location = new System.Drawing.Point(12, 12);
            this.workProgress.Name = "workProgress";
            this.workProgress.Size = new System.Drawing.Size(446, 23);
            this.workProgress.TabIndex = 0;
            // 
            // ProgBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 47);
            this.Controls.Add(this.workProgress);
            this.Name = "ProgBar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Analyzing Spreadsheet...";
            this.Load += new System.EventHandler(this.ProgBar_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar workProgress;
    }
}