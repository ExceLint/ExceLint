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
            this.overwriteButton = new System.Windows.Forms.Button();
            this.appendButton = new System.Windows.Forms.Button();
            this.messageText = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // overwriteButton
            // 
            this.overwriteButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.overwriteButton.Location = new System.Drawing.Point(24, 126);
            this.overwriteButton.Name = "overwriteButton";
            this.overwriteButton.Size = new System.Drawing.Size(105, 23);
            this.overwriteButton.TabIndex = 1;
            this.overwriteButton.Text = "Overwrite";
            this.overwriteButton.UseVisualStyleBackColor = true;
            // 
            // appendButton
            // 
            this.appendButton.Location = new System.Drawing.Point(151, 126);
            this.appendButton.Name = "appendButton";
            this.appendButton.Size = new System.Drawing.Size(111, 23);
            this.appendButton.TabIndex = 2;
            this.appendButton.Text = "Append";
            this.appendButton.UseVisualStyleBackColor = true;
            // 
            // messageText
            // 
            this.messageText.AutoSize = true;
            this.messageText.Location = new System.Drawing.Point(21, 20);
            this.messageText.MaximumSize = new System.Drawing.Size(80, 0);
            this.messageText.Name = "messageText";
            this.messageText.Size = new System.Drawing.Size(35, 13);
            this.messageText.TabIndex = 3;
            this.messageText.Text = "label1";
            // 
            // OverwriteOrAppend
            // 
            this.AcceptButton = this.appendButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.overwriteButton;
            this.ClientSize = new System.Drawing.Size(284, 161);
            this.Controls.Add(this.messageText);
            this.Controls.Add(this.appendButton);
            this.Controls.Add(this.overwriteButton);
            this.Name = "OverwriteOrAppend";
            this.Text = "OverwriteOrAppend";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button overwriteButton;
        private System.Windows.Forms.Button appendButton;
        private System.Windows.Forms.Label messageText;
    }
}