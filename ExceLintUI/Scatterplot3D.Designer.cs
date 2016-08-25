namespace ExceLintUI
{
    partial class Scatterplot3D
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.labelCondition = new System.Windows.Forms.Label();
            this.comboCondition = new System.Windows.Forms.ComboBox();
            this.labelFeature = new System.Windows.Forms.Label();
            this.comboFeature = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // chart1
            // 
            chartArea3.Area3DStyle.Enable3D = true;
            chartArea3.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea3);
            legend3.Name = "Legend1";
            this.chart1.Legends.Add(legend3);
            this.chart1.Location = new System.Drawing.Point(52, 27);
            this.chart1.Name = "chart1";
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series3.Legend = "Legend1";
            series3.Name = "Series1";
            this.chart1.Series.Add(series3);
            this.chart1.Size = new System.Drawing.Size(658, 420);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "scatterplot";
            // 
            // labelCondition
            // 
            this.labelCondition.AutoSize = true;
            this.labelCondition.Location = new System.Drawing.Point(382, 472);
            this.labelCondition.Name = "labelCondition";
            this.labelCondition.Size = new System.Drawing.Size(54, 13);
            this.labelCondition.TabIndex = 8;
            this.labelCondition.Text = "Condition:";
            this.labelCondition.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboCondition
            // 
            this.comboCondition.FormattingEnabled = true;
            this.comboCondition.Location = new System.Drawing.Point(434, 469);
            this.comboCondition.Name = "comboCondition";
            this.comboCondition.Size = new System.Drawing.Size(221, 21);
            this.comboCondition.TabIndex = 7;
            // 
            // labelFeature
            // 
            this.labelFeature.AutoSize = true;
            this.labelFeature.Location = new System.Drawing.Point(77, 472);
            this.labelFeature.Name = "labelFeature";
            this.labelFeature.Size = new System.Drawing.Size(46, 13);
            this.labelFeature.TabIndex = 6;
            this.labelFeature.Text = "Feature:";
            this.labelFeature.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboFeature
            // 
            this.comboFeature.FormattingEnabled = true;
            this.comboFeature.Location = new System.Drawing.Point(129, 469);
            this.comboFeature.Name = "comboFeature";
            this.comboFeature.Size = new System.Drawing.Size(221, 21);
            this.comboFeature.TabIndex = 5;
            // 
            // Scaterplot3D
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(765, 502);
            this.Controls.Add(this.labelCondition);
            this.Controls.Add(this.comboCondition);
            this.Controls.Add(this.labelFeature);
            this.Controls.Add(this.comboFeature);
            this.Controls.Add(this.chart1);
            this.Name = "Scaterplot3D";
            this.Text = "Scaterplot3D";
            this.Load += new System.EventHandler(this.Scaterplot3D_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Label labelCondition;
        private System.Windows.Forms.ComboBox comboCondition;
        private System.Windows.Forms.Label labelFeature;
        private System.Windows.Forms.ComboBox comboFeature;
    }
}