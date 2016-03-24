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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.CheckCellGroup = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.AnalyzeButton = this.Factory.CreateRibbonButton();
            this.MarkAsOKButton = this.Factory.CreateRibbonButton();
            this.FixErrorButton = this.Factory.CreateRibbonButton();
            this.StartOverButton = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.FrmAbsVect = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.DataAbsVect = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.colSelect = this.Factory.CreateRibbonButton();
            this.rowSelected = this.Factory.CreateRibbonButton();
            this.ToDOT = this.Factory.CreateRibbonButton();
            this.significanceTextBox = this.Factory.CreateRibbonEditBox();
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
            this.CheckCellGroup.Name = "CheckCellGroup";
            // 
            // box1
            // 
            this.box1.Items.Add(this.AnalyzeButton);
            this.box1.Items.Add(this.MarkAsOKButton);
            this.box1.Items.Add(this.FixErrorButton);
            this.box1.Items.Add(this.StartOverButton);
            this.box1.Items.Add(this.button1);
            this.box1.Items.Add(this.FrmAbsVect);
            this.box1.Items.Add(this.button3);
            this.box1.Items.Add(this.DataAbsVect);
            this.box1.Items.Add(this.button2);
            this.box1.Items.Add(this.colSelect);
            this.box1.Items.Add(this.rowSelected);
            this.box1.Items.Add(this.ToDOT);
            this.box1.Name = "box1";
            // 
            // AnalyzeButton
            // 
            this.AnalyzeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AnalyzeButton.Image = global::ExceLintUI.Properties.Resources.analyze_small;
            this.AnalyzeButton.Label = "Analyze";
            this.AnalyzeButton.Name = "AnalyzeButton";
            this.AnalyzeButton.ShowImage = true;
            this.AnalyzeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AnalyzeButton_Click);
            // 
            // MarkAsOKButton
            // 
            this.MarkAsOKButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MarkAsOKButton.Image = global::ExceLintUI.Properties.Resources.mark_as_ok_small;
            this.MarkAsOKButton.Label = "Mark as OK";
            this.MarkAsOKButton.Name = "MarkAsOKButton";
            this.MarkAsOKButton.ShowImage = true;
            this.MarkAsOKButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MarkAsOKButton_Click);
            // 
            // FixErrorButton
            // 
            this.FixErrorButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FixErrorButton.Image = global::ExceLintUI.Properties.Resources.correct_small;
            this.FixErrorButton.Label = "Fix Error";
            this.FixErrorButton.Name = "FixErrorButton";
            this.FixErrorButton.ShowImage = true;
            this.FixErrorButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FixErrorButton_Click);
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
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::ExceLintUI.Properties.Resources.pain;
            this.button1.Label = "Input (rel)";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Visible = false;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // FrmAbsVect
            // 
            this.FrmAbsVect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FrmAbsVect.Image = global::ExceLintUI.Properties.Resources.pain;
            this.FrmAbsVect.Label = "Input (abs)";
            this.FrmAbsVect.Name = "FrmAbsVect";
            this.FrmAbsVect.ShowImage = true;
            this.FrmAbsVect.Visible = false;
            this.FrmAbsVect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FrmAbsVect_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::ExceLintUI.Properties.Resources.pain;
            this.button3.Label = "Output (rel)";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Visible = false;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // DataAbsVect
            // 
            this.DataAbsVect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DataAbsVect.Image = global::ExceLintUI.Properties.Resources.pain;
            this.DataAbsVect.Label = "Output (abs)";
            this.DataAbsVect.Name = "DataAbsVect";
            this.DataAbsVect.ShowImage = true;
            this.DataAbsVect.Visible = false;
            this.DataAbsVect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DataAbsVect_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::ExceLintUI.Properties.Resources.pain;
            this.button2.Label = "L2Sum";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Visible = false;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // colSelect
            // 
            this.colSelect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.colSelect.Image = global::ExceLintUI.Properties.Resources.pain;
            this.colSelect.Label = "ColSel";
            this.colSelect.Name = "colSelect";
            this.colSelect.ShowImage = true;
            this.colSelect.Visible = false;
            this.colSelect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colSelect_Click);
            // 
            // rowSelected
            // 
            this.rowSelected.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.rowSelected.Image = global::ExceLintUI.Properties.Resources.pain;
            this.rowSelected.Label = "RowSel";
            this.rowSelected.Name = "rowSelected";
            this.rowSelected.ShowImage = true;
            this.rowSelected.Visible = false;
            this.rowSelected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rowSelected_Click);
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
            // significanceTextBox
            // 
            this.significanceTextBox.Label = "Sig. Thresh.";
            this.significanceTextBox.Name = "significanceTextBox";
            this.significanceTextBox.SizeString = "100.0";
            this.significanceTextBox.Text = "0.25";
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
            // 
            // outDegree
            // 
            this.outDegree.Label = "Out-Degree";
            this.outDegree.Name = "outDegree";
            this.outDegree.Visible = false;
            // 
            // combinedDegree
            // 
            this.combinedDegree.Label = "Both-Degree";
            this.combinedDegree.Name = "combinedDegree";
            this.combinedDegree.Visible = false;
            // 
            // inVectors
            // 
            this.inVectors.Checked = true;
            this.inVectors.Label = "In-Vectors (mix; sh)";
            this.inVectors.Name = "inVectors";
            this.inVectors.Visible = false;
            // 
            // outVectors
            // 
            this.outVectors.Label = "Out-Vectors (mix; sh)";
            this.outVectors.Name = "outVectors";
            this.outVectors.Visible = false;
            // 
            // inVectorsAbs
            // 
            this.inVectorsAbs.Label = "In-Vectors (abs; sh)";
            this.inVectorsAbs.Name = "inVectorsAbs";
            this.inVectorsAbs.Visible = false;
            // 
            // outVectorsAbs
            // 
            this.outVectorsAbs.Label = "Out-Vectors (abs; sh)";
            this.outVectorsAbs.Name = "outVectorsAbs";
            this.outVectorsAbs.Visible = false;
            // 
            // ProximityAbove
            // 
            this.ProximityAbove.Label = "Above";
            this.ProximityAbove.Name = "ProximityAbove";
            this.ProximityAbove.Visible = false;
            // 
            // ProximityBelow
            // 
            this.ProximityBelow.Label = "Below";
            this.ProximityBelow.Name = "ProximityBelow";
            this.ProximityBelow.Visible = false;
            // 
            // ProximityLeft
            // 
            this.ProximityLeft.Label = "Left";
            this.ProximityLeft.Name = "ProximityLeft";
            this.ProximityLeft.Visible = false;
            // 
            // ProximityRight
            // 
            this.ProximityRight.Label = "Right";
            this.ProximityRight.Name = "ProximityRight";
            this.ProximityRight.Visible = false;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // allCellsFreq
            // 
            this.allCellsFreq.Checked = true;
            this.allCellsFreq.Label = "All Cells Freq.";
            this.allCellsFreq.Name = "allCellsFreq";
            // 
            // columnCellsFreq
            // 
            this.columnCellsFreq.Checked = true;
            this.columnCellsFreq.Label = "Column Cells Freq";
            this.columnCellsFreq.Name = "columnCellsFreq";
            // 
            // rowCellsFreq
            // 
            this.rowCellsFreq.Checked = true;
            this.rowCellsFreq.Label = "Row Cells Freq";
            this.rowCellsFreq.Name = "rowCellsFreq";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FixErrorButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StartOverButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rowSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FrmAbsVect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DataAbsVect;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityAbove;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityBelow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityLeft;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox ProximityRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox significanceTextBox;
    }

    partial class ThisRibbonCollection
    {
        internal ExceLintRibbon ExceLintRibbon
        {
            get { return this.GetRibbon<ExceLintRibbon>(); }
        }
    }
}
