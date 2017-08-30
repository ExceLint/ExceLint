namespace ExceLintUI
{
    partial class OverwriteOrAppend
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
            this.messageText = new System.Windows.Forms.Label();
            this.appendButton = new System.Windows.Forms.Button();
            this.overwriteButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // messageText
            // 
            this.messageText.AutoSize = true;
            this.messageText.Location = new System.Drawing.Point(21, 20);
            this.messageText.MaximumSize = new System.Drawing.Size(260, 50);
            this.messageText.Name = "messageText";
            this.messageText.Size = new System.Drawing.Size(29, 13);
            this.messageText.TabIndex = 3;
            this.messageText.Text = "label";
            // 
            // appendButton
            // 
            this.appendButton.Location = new System.Drawing.Point(197, 66);
            this.appendButton.Name = "appendButton";
            this.appendButton.Size = new System.Drawing.Size(75, 23);
            this.appendButton.TabIndex = 4;
            this.appendButton.Text = "Append";
            this.appendButton.UseVisualStyleBackColor = true;
            this.appendButton.Click += new System.EventHandler(this.appendButton_Click_1);
            // 
            // overwriteButton
            // 
            this.overwriteButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.overwriteButton.Location = new System.Drawing.Point(116, 66);
            this.overwriteButton.Name = "overwriteButton";
            this.overwriteButton.Size = new System.Drawing.Size(75, 23);
            this.overwriteButton.TabIndex = 5;
            this.overwriteButton.Text = "Overwrite";
            this.overwriteButton.UseVisualStyleBackColor = true;
            this.overwriteButton.Click += new System.EventHandler(this.overwriteButton_Click_1);
            // 
            // OverwriteOrAppend
            // 
            this.AcceptButton = this.appendButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.overwriteButton;
            this.ClientSize = new System.Drawing.Size(284, 101);
            this.Controls.Add(this.overwriteButton);
            this.Controls.Add(this.appendButton);
            this.Controls.Add(this.messageText);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.Name = "OverwriteOrAppend";
            this.Text = "Overwrite or append?";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label messageText;
        private System.Windows.Forms.Button appendButton;
        private System.Windows.Forms.Button overwriteButton;
    }
}