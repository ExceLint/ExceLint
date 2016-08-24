namespace ExceLintUI
{
    partial class SpectralPlot
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.comboFeature = new System.Windows.Forms.ComboBox();
            this.labelFeature = new System.Windows.Forms.Label();
            this.labelCondition = new System.Windows.Forms.Label();
            this.comboCondition = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(12, 12);
            this.chart1.Name = "chart1";
            this.chart1.Size = new System.Drawing.Size(787, 545);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "chart1";
            this.chart1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.chart1_MouseMove);
            // 
            // comboFeature
            // 
            this.comboFeature.FormattingEnabled = true;
            this.comboFeature.Location = new System.Drawing.Point(65, 563);
            this.comboFeature.Name = "comboFeature";
            this.comboFeature.Size = new System.Drawing.Size(221, 21);
            this.comboFeature.TabIndex = 1;
            // 
            // labelFeature
            // 
            this.labelFeature.AutoSize = true;
            this.labelFeature.Location = new System.Drawing.Point(13, 566);
            this.labelFeature.Name = "labelFeature";
            this.labelFeature.Size = new System.Drawing.Size(46, 13);
            this.labelFeature.TabIndex = 2;
            this.labelFeature.Text = "Feature:";
            this.labelFeature.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelCondition
            // 
            this.labelCondition.AutoSize = true;
            this.labelCondition.Location = new System.Drawing.Point(318, 566);
            this.labelCondition.Name = "labelCondition";
            this.labelCondition.Size = new System.Drawing.Size(54, 13);
            this.labelCondition.TabIndex = 4;
            this.labelCondition.Text = "Condition:";
            this.labelCondition.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboCondition
            // 
            this.comboCondition.FormattingEnabled = true;
            this.comboCondition.Location = new System.Drawing.Point(370, 563);
            this.comboCondition.Name = "comboCondition";
            this.comboCondition.Size = new System.Drawing.Size(221, 21);
            this.comboCondition.TabIndex = 3;
            this.comboCondition.SelectedIndexChanged += new System.EventHandler(this.comboCondition_SelectedIndexChanged);
            // 
            // SpectralPlot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(811, 596);
            this.Controls.Add(this.labelCondition);
            this.Controls.Add(this.comboCondition);
            this.Controls.Add(this.labelFeature);
            this.Controls.Add(this.comboFeature);
            this.Controls.Add(this.chart1);
            this.Name = "SpectralPlot";
            this.Text = "SpectralPlot";
            this.Load += new System.EventHandler(this.SpectralPlot_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.ComboBox comboFeature;
        private System.Windows.Forms.Label labelFeature;
        private System.Windows.Forms.Label labelCondition;
        private System.Windows.Forms.ComboBox comboCondition;
    }
}