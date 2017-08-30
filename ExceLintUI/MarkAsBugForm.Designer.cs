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
            this.editNotes = new System.Windows.Forms.CheckBox();
            this.editBugKind = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BugKindsCombo
            // 
            this.BugKindsCombo.FormattingEnabled = true;
            this.BugKindsCombo.Location = new System.Drawing.Point(15, 32);
            this.BugKindsCombo.Name = "BugKindsCombo";
            this.BugKindsCombo.Size = new System.Drawing.Size(360, 21);
            this.BugKindsCombo.TabIndex = 0;
            // 
            // markCellLabel
            // 
            this.markCellLabel.AutoSize = true;
            this.markCellLabel.Location = new System.Drawing.Point(12, 16);
            this.markCellLabel.Name = "markCellLabel";
            this.markCellLabel.Size = new System.Drawing.Size(92, 13);
            this.markCellLabel.TabIndex = 1;
            this.markCellLabel.Text = "Mark cell FOO as:";
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(219, 146);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // bugNotesTextField
            // 
            this.bugNotesTextField.Location = new System.Drawing.Point(15, 83);
            this.bugNotesTextField.Multiline = true;
            this.bugNotesTextField.Name = "bugNotesTextField";
            this.bugNotesTextField.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.bugNotesTextField.Size = new System.Drawing.Size(360, 57);
            this.bugNotesTextField.TabIndex = 3;
            // 
            // markButton
            // 
            this.markButton.Location = new System.Drawing.Point(300, 146);
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
            this.label2.Location = new System.Drawing.Point(12, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Notes:";
            // 
            // editNotes
            // 
            this.editNotes.AutoSize = true;
            this.editNotes.Location = new System.Drawing.Point(396, 83);
            this.editNotes.Name = "editNotes";
            this.editNotes.Size = new System.Drawing.Size(15, 14);
            this.editNotes.TabIndex = 8;
            this.editNotes.UseVisualStyleBackColor = true;
            // 
            // editBugKind
            // 
            this.editBugKind.AutoSize = true;
            this.editBugKind.Location = new System.Drawing.Point(396, 35);
            this.editBugKind.Name = "editBugKind";
            this.editBugKind.Size = new System.Drawing.Size(15, 14);
            this.editBugKind.TabIndex = 6;
            this.editBugKind.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(383, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Edit All";
            // 
            // MarkAsBugForm
            // 
            this.AcceptButton = this.markButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(438, 181);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.editNotes);
            this.Controls.Add(this.editBugKind);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.markButton);
            this.Controls.Add(this.bugNotesTextField);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.markCellLabel);
            this.Controls.Add(this.BugKindsCombo);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "MarkAsBugForm";
            this.Text = "Annotate cell";
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
        private System.Windows.Forms.CheckBox editNotes;
        private System.Windows.Forms.CheckBox editBugKind;
        private System.Windows.Forms.Label label1;
    }
}