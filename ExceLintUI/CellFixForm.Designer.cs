namespace ExceLintUI
{
    partial class CellFixForm
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
            this.FixText = new System.Windows.Forms.TextBox();
            this.FixButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // FixText
            // 
            this.FixText.Location = new System.Drawing.Point(12, 22);
            this.FixText.Name = "FixText";
            this.FixText.Size = new System.Drawing.Size(156, 20);
            this.FixText.TabIndex = 0;
            // 
            // FixButton
            // 
            this.FixButton.Location = new System.Drawing.Point(12, 54);
            this.FixButton.Name = "FixButton";
            this.FixButton.Size = new System.Drawing.Size(75, 23);
            this.FixButton.TabIndex = 1;
            this.FixButton.Text = "Fix";
            this.FixButton.UseVisualStyleBackColor = true;
            this.FixButton.Click += new System.EventHandler(this.FixButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(93, 54);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 2;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // CellFixForm
            // 
            this.AcceptButton = this.FixButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(179, 95);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.FixButton);
            this.Controls.Add(this.FixText);
            this.Name = "CellFixForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "CellFixForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox FixText;
        private System.Windows.Forms.Button FixButton;
        private System.Windows.Forms.Button CancelButton;
    }
}