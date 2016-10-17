namespace ExceLintUI
{
    partial class MarkAsBugForm
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
            this.BugKindsCombo = new System.Windows.Forms.ComboBox();
            this.markCellLabel = new System.Windows.Forms.Label();
            this.cancelButton = new System.Windows.Forms.Button();
            this.bugNotesTextField = new System.Windows.Forms.TextBox();
            this.markButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BugKindsCombo
            // 
            this.BugKindsCombo.FormattingEnabled = true;
            this.BugKindsCombo.Location = new System.Drawing.Point(110, 6);
            this.BugKindsCombo.Name = "BugKindsCombo";
            this.BugKindsCombo.Size = new System.Drawing.Size(265, 21);
            this.BugKindsCombo.TabIndex = 0;
            // 
            // markCellLabel
            // 
            this.markCellLabel.AutoSize = true;
            this.markCellLabel.Location = new System.Drawing.Point(12, 9);
            this.markCellLabel.Name = "markCellLabel";
            this.markCellLabel.Size = new System.Drawing.Size(92, 13);
            this.markCellLabel.TabIndex = 1;
            this.markCellLabel.Text = "Mark cell FOO as:";
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(219, 112);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // bugNotesCombo
            // 
            this.bugNotesTextField.Location = new System.Drawing.Point(15, 49);
            this.bugNotesTextField.Multiline = true;
            this.bugNotesTextField.Name = "bugNotesCombo";
            this.bugNotesTextField.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.bugNotesTextField.Size = new System.Drawing.Size(360, 57);
            this.bugNotesTextField.TabIndex = 3;
            // 
            // markButton
            // 
            this.markButton.Location = new System.Drawing.Point(300, 112);
            this.markButton.Name = "markButton";
            this.markButton.Size = new System.Drawing.Size(75, 23);
            this.markButton.TabIndex = 4;
            this.markButton.Text = "Mark";
            this.markButton.UseVisualStyleBackColor = true;
            this.markButton.Click += new System.EventHandler(this.markButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Notes:";
            // 
            // MarkAsBugForm
            // 
            this.AcceptButton = this.markButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(387, 147);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.markButton);
            this.Controls.Add(this.bugNotesTextField);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.markCellLabel);
            this.Controls.Add(this.BugKindsCombo);
            this.Name = "MarkAsBugForm";
            this.Text = "MarkAsBugForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox BugKindsCombo;
        private System.Windows.Forms.Label markCellLabel;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.TextBox bugNotesTextField;
        private System.Windows.Forms.Button markButton;
        private System.Windows.Forms.Label label2;
    }
}