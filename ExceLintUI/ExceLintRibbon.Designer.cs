namespace ExceLintUI
{
    partial class ExceLintRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExceLintRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExceLintRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.CheckCellGroup = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.AnalyzeButton = this.Factory.CreateRibbonButton();
            this.MarkAsOKButton = this.Factory.CreateRibbonButton();
            this.StartOverButton = this.Factory.CreateRibbonButton();
            this.showHeatmap = this.Factory.CreateRibbonButton();
            this.ToDOT = this.Factory.CreateRibbonButton();
            this.showVectors = this.Factory.CreateRibbonButton();
            this.spectralPlot = this.Factory.CreateRibbonButton();
            this.scatter3D = this.Factory.CreateRibbonButton();
            this.significanceTextBox = this.Factory.CreateRibbonEditBox();
            this.spectralRanking = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.inDegree = this.Factory.CreateRibbonCheckBox();
            this.outDegree = this.Factory.CreateRibbonCheckBox();
            this.combinedDegree = this.Factory.CreateRibbonCheckBox();
            this.inVectors = this.Factory.CreateRibbonCheckBox();
            this.outVectors = this.Factory.CreateRibbonCheckBox();
            this.inVectorsAbs = this.Factory.CreateRibbonCheckBox();
            this.outVectorsAbs = this.Factory.CreateRibbonCheckBox();
            this.ProximityAbove = this.Factory.CreateRibbonCheckBox();
            this.ProximityBelow = this.Factory.CreateRibbonCheckBox();
            this.ProximityLeft = this.Factory.CreateRibbonCheckBox();
            this.ProximityRight = this.Factory.CreateRibbonCheckBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.allCellsFreq = this.Factory.CreateRibbonCheckBox();
            this.columnCellsFreq = this.Factory.CreateRibbonCheckBox();
            this.rowCellsFreq = this.Factory.CreateRibbonCheckBox();
            this.levelsFreq = this.Factory.CreateRibbonCheckBox();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.DebugOutput = this.Factory.CreateRibbonCheckBox();
            this.forceBuildDAG = this.Factory.CreateRibbonCheckBox();
            this.inferAddrModes = this.Factory.CreateRibbonCheckBox();
            this.allCells = this.Factory.CreateRibbonCheckBox();
            this.weightByIntrinsicAnomalousness = this.Factory.CreateRibbonCheckBox();
            this.conditioningSetSize = this.Factory.CreateRibbonCheckBox();
            this.showFixes = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.CheckCellGroup.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.CheckCellGroup);
            this.tab2.Label = "ExceLint";
            this.tab2.Name = "tab2";
            // 
            // CheckCellGroup
            // 
            this.CheckCellGroup.Items.Add(this.box1);
            this.CheckCellGroup.Items.Add(this.significanceTextBox);
            this.CheckCellGroup.Items.Add(this.spectralRanking);
            this.CheckCellGroup.Items.Add(this.showFixes);
            this.CheckCellGroup.Items.Add(this.separator1);
            this.CheckCellGroup.Items.Add(this.inDegree);
            this.CheckCellGroup.Items.Add(this.outDegree);
            this.CheckCellGroup.Items.Add(this.combinedDegree);
            this.CheckCellGroup.Items.Add(this.inVectors);
            this.CheckCellGroup.Items.Add(this.outVectors);
            this.CheckCellGroup.Items.Add(this.inVectorsAbs);
            this.CheckCellGroup.Items.Add(this.outVectorsAbs);
            this.CheckCellGroup.Items.Add(this.ProximityAbove);
            this.CheckCellGroup.Items.Add(this.ProximityBelow);
            this.CheckCellGroup.Items.Add(this.ProximityLeft);
            this.CheckCellGroup.Items.Add(this.ProximityRight);
            this.CheckCellGroup.Items.Add(this.separator2);
            this.CheckCellGroup.Items.Add(this.allCellsFreq);
            this.CheckCellGroup.Items.Add(this.columnCellsFreq);
            this.CheckCellGroup.Items.Add(this.rowCellsFreq);
            this.CheckCellGroup.Items.Add(this.levelsFreq);
            this.CheckCellGroup.Items.Add(this.separator3);
            this.CheckCellGroup.Items.Add(this.DebugOutput);
            this.CheckCellGroup.Items.Add(this.forceBuildDAG);
            this.CheckCellGroup.Items.Add(this.inferAddrModes);
            this.CheckCellGroup.Items.Add(this.allCells);
            this.CheckCellGroup.Items.Add(this.weightByIntrinsicAnomalousness);
            this.CheckCellGroup.Items.Add(this.conditioningSetSize);
            this.CheckCellGroup.Name = "CheckCellGroup";
            // 
            // box1
            // 
            this.box1.Items.Add(this.AnalyzeButton);
            this.box1.Items.Add(this.MarkAsOKButton);
            this.box1.Items.Add(this.StartOverButton);
            this.box1.Items.Add(this.showHeatmap);
            this.box1.Items.Add(this.ToDOT);
            this.box1.Items.Add(this.showVectors);
            this.box1.Items.Add(this.spectralPlot);
            this.box1.Items.Add(this.scatter3D);
            this.box1.Name = "box1";
            // 
            // AnalyzeButton
            // 
            this.AnalyzeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AnalyzeButton.Image = global::ExceLintUI.Properties.Resources.analyze_small;
            this.AnalyzeButton.Label = "Audit";
            this.AnalyzeButton.Name = "AnalyzeButton";
            this.AnalyzeButton.ShowImage = true;
            this.AnalyzeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AnalyzeButton_Click);
            // 
            // MarkAsOKButton
            // 
            this.MarkAsOKButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MarkAsOKButton.Image = global::ExceLintUI.Properties.Resources.mark_as_ok_small;
            this.MarkAsOKButton.Label = "Next Cell";
            this.MarkAsOKButton.Name = "MarkAsOKButton";
            this.MarkAsOKButton.ShowImage = true;
            this.MarkAsOKButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MarkAsOKButton_Click);
            // 
            // StartOverButton
            // 
            this.StartOverButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StartOverButton.Image = global::ExceLintUI.Properties.Resources.clear_small;
            this.StartOverButton.Label = "Start Over";
            this.StartOverButton.Name = "StartOverButton";
            this.StartOverButton.ShowImage = true;
            this.StartOverButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartOverButton_Click);
            // 
            // showHeatmap
            // 
            this.showHeatmap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showHeatmap.Image = ((System.Drawing.Image)(resources.GetObject("showHeatmap.Image")));
            this.showHeatmap.Label = "Heat Map";
            this.showHeatmap.Name = "showHeatmap";
            this.showHeatmap.ShowImage = true;
            this.showHeatmap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showHeatmap_Click);
            // 
            // ToDOT
            // 
            this.ToDOT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ToDOT.Image = global::ExceLintUI.Properties.Resources.graph;
            this.ToDOT.Label = "ToDOT";
            this.ToDOT.Name = "ToDOT";
            this.ToDOT.ShowImage = true;
            this.ToDOT.Visible = false;
            this.ToDOT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToDOT_Click);
            // 
            // showVectors
            // 
            this.showVectors.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showVectors.Image = global::ExceLintUI.Properties.Resources.graph;
            this.showVectors.Label = "Show Vectors";
            this.showVectors.Name = "showVectors";
            this.showVectors.ShowImage = true;
            this.showVectors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showVectors_Click);
            // 
            // spectralPlot
            // 
            this.spectralPlot.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.spectralPlot.Image = global::ExceLintUI.Properties.Resources.spectral_plot_32;
            this.spectralPlot.Label = "Spectral Plot";
            this.spectralPlot.Name = "spectralPlot";
            this.spectralPlot.ShowImage = true;
            this.spectralPlot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spectralPlot_Click);
            // 
            // scatter3D
            // 
            this.scatter3D.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.scatter3D.Image = global::ExceLintUI.Properties.Resources.spectral_plot_32;
            this.scatter3D.Label = "3D Scatterplot";
            this.scatter3D.Name = "scatter3D";
            this.scatter3D.ShowImage = true;
            this.scatter3D.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.scatter3D_Click);
            // 
            // significanceTextBox
            // 
            this.significanceTextBox.Label = "Error Rate %";
            this.significanceTextBox.Name = "significanceTextBox";
            this.significanceTextBox.SizeString = "100.0";
            this.significanceTextBox.Text = "5";
            this.significanceTextBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.significanceTextBox_TextChanged);
            // 
            // spectralRanking
            // 
            this.spectralRanking.Label = "Use Spectral Rank";
            this.spectralRanking.Name = "spectralRanking";
            this.spectralRanking.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spectralRanking_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // inDegree
            // 
            this.inDegree.Label = "In-Degree";
            this.inDegree.Name = "inDegree";
            this.inDegree.Visible = false;
            this.inDegree.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inDegree_Click);
            // 
            // outDegree
            // 
            this.outDegree.Label = "Out-Degree";
            this.outDegree.Name = "outDegree";
            this.outDegree.Visible = false;
            this.outDegree.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.outDegree_Click);
            // 
            // combinedDegree
            // 
            this.combinedDegree.Label = "Both-Degree";
            this.combinedDegree.Name = "combinedDegree";
            this.combinedDegree.Visible = false;
            this.combinedDegree.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.combinedDegree_Click);
            // 
            // inVectors
            // 
            this.inVectors.Checked = true;
            this.inVectors.Label = "In-Vectors (mix; sh)";
            this.inVectors.Name = "inVectors";
            this.inVectors.Visible = false;
            this.inVectors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inVectors_Click);
            // 
            // outVectors
            // 
            this.outVectors.Label = "Out-Vectors (mix; sh)";
            this.outVectors.Name = "outVectors";
            this.outVectors.Visible = false;
            this.outVectors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.outVectors_Click);
            // 
            // inVectorsAbs
            // 
            this.inVectorsAbs.Label = "In-Vectors (abs; sh)";
            this.inVectorsAbs.Name = "inVectorsAbs";
            this.inVectorsAbs.Visible = false;
            this.inVectorsAbs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inVectorsAbs_Click);
            // 
            // outVectorsAbs
            // 
            this.outVectorsAbs.Label = "Out-Vectors (abs; sh)";
            this.outVectorsAbs.Name = "outVectorsAbs";
            this.outVectorsAbs.Visible = false;
            this.outVectorsAbs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.outVectorsAbs_Click);
            // 
            // ProximityAbove
            // 
            this.ProximityAbove.Label = "Above";
            this.ProximityAbove.Name = "ProximityAbove";
            this.ProximityAbove.Visible = false;
            this.ProximityAbove.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProximityAbove_Click);
            // 
            // ProximityBelow
            // 
            this.ProximityBelow.Label = "Below";
            this.ProximityBelow.Name = "ProximityBelow";
            this.ProximityBelow.Visible = false;
            this.ProximityBelow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProximityBelow_Click);
            // 
            // ProximityLeft
            // 
            this.ProximityLeft.Label = "Left";
            this.ProximityLeft.Name = "ProximityLeft";
            this.ProximityLeft.Visible = false;
            this.ProximityLeft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProximityLeft_Click);
            // 
            // ProximityRight
            // 
            this.ProximityRight.Label = "Right";
            this.ProximityRight.Name = "ProximityRight";
            this.ProximityRight.Visible = false;
            this.ProximityRight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProximityRight_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // allCellsFreq
            // 
            this.allCellsFreq.Checked = true;
            this.allCellsFreq.Label = "All Cells Freq";
            this.allCellsFreq.Name = "allCellsFreq";
            this.allCellsFreq.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.allCellsFreq_Click);
            // 
            // columnCellsFreq
            // 
            this.columnCellsFreq.Checked = true;
            this.columnCellsFreq.Label = "Column Cells Freq";
            this.columnCellsFreq.Name = "columnCellsFreq";
            this.columnCellsFreq.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.columnCellsFreq_Click);
            // 
            // rowCellsFreq
            // 
            this.rowCellsFreq.Checked = true;
            this.rowCellsFreq.Label = "Row Cells Freq";
            this.rowCellsFreq.Name = "rowCellsFreq";
            this.rowCellsFreq.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rowCellsFreq_Click);
            // 
            // levelsFreq
            // 
            this.levelsFreq.Checked = true;
            this.levelsFreq.Label = "Levels Freq";
            this.levelsFreq.Name = "levelsFreq";
            this.levelsFreq.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.levelsFreq_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // DebugOutput
            // 
            this.DebugOutput.Label = "Show Debug Output";
            this.DebugOutput.Name = "DebugOutput";
            this.DebugOutput.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DebugOutput_Click);
            // 
            // forceBuildDAG
            // 
            this.forceBuildDAG.Label = "Force DAG Rebuild";
            this.forceBuildDAG.Name = "forceBuildDAG";
            this.forceBuildDAG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.forceBuildDAG_Click);
            // 
            // inferAddrModes
            // 
            this.inferAddrModes.Label = "Infer Address Modes";
            this.inferAddrModes.Name = "inferAddrModes";
            this.inferAddrModes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inferAddrModes_Click);
            // 
            // allCells
            // 
            this.allCells.Label = "Analyze All Cells";
            this.allCells.Name = "allCells";
            this.allCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.allCells_Click);
            // 
            // weightByIntrinsicAnomalousness
            // 
            this.weightByIntrinsicAnomalousness.Label = "Reweight by Intrinsic Anomalousness";
            this.weightByIntrinsicAnomalousness.Name = "weightByIntrinsicAnomalousness";
            this.weightByIntrinsicAnomalousness.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.weightByIntrinsicAnomalousness_Click);
            // 
            // conditioningSetSize
            // 
            this.conditioningSetSize.Checked = true;
            this.conditioningSetSize.Label = "Weigh by Conditioning Set Size";
            this.conditioningSetSize.Name = "conditioningSetSize";
            this.conditioningSetSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.conditioningSetSize_Click);
            // 
            // showFixes
            // 
            this.showFixes.Enabled = false;
            this.showFixes.Label = "Show Fixes";
            this.showFixes.Name = "showFixes";
            // 
            // ExceLintRibbon
            // 
            this.Name = "ExceLintRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExceLintRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.CheckCellGroup.ResumeLayout(false);
            this.CheckCellGroup.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CheckCellGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AnalyzeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkAsOKButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StartOverButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox inDegree;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox outDegree;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox combinedDegree;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox inVectors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ToDOT;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox outVectors;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox inVectorsAbs;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox outVectorsAbs;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox allCellsFreq;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox columnCellsFreq;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rowCellsFreq;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityAbove;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityBelow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox significanceTextBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox DebugOutput;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showHeatmap;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox forceBuildDAG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showVectors;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox inferAddrModes;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox allCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox weightByIntrinsicAnomalousness;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox levelsFreq;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox conditioningSetSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton spectralPlot;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox spectralRanking;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton scatter3D;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox showFixes;
    }

    partial class ThisRibbonCollection
    {
        internal ExceLintRibbon ExceLintRibbon
        {
            get { return this.GetRibbon<ExceLintRibbon>(); }
        }
    }
}
