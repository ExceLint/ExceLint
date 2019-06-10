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
            this.VectorForCell = this.Factory.CreateRibbonButton();
            this.analyzeFormulas = this.Factory.CreateRibbonCheckBox();
            this.enableDataHighlight = this.Factory.CreateRibbonCheckBox();
            this.ComputeEntropy = this.Factory.CreateRibbonButton();
            this.RegularityMap = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.CheckCellGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            this.tab1.Visible = false;
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.CheckCellGroup);
            this.tab2.Label = "ExceLint";
            this.tab2.Name = "tab2";
            // 
            // CheckCellGroup
            // 
            this.CheckCellGroup.Items.Add(this.RegularityMap);
            this.CheckCellGroup.Items.Add(this.VectorForCell);
            this.CheckCellGroup.Items.Add(this.analyzeFormulas);
            this.CheckCellGroup.Items.Add(this.enableDataHighlight);
            this.CheckCellGroup.Items.Add(this.ComputeEntropy);
            this.CheckCellGroup.Name = "CheckCellGroup";
            // 
            // VectorForCell
            // 
            this.VectorForCell.Label = "";
            this.VectorForCell.Name = "VectorForCell";
            // 
            // analyzeFormulas
            // 
            this.analyzeFormulas.Label = "Analyze formulas";
            this.analyzeFormulas.Name = "analyzeFormulas";
            this.analyzeFormulas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.analyzeFormulas_Click);
            // 
            // enableDataHighlight
            // 
            this.enableDataHighlight.Label = "Analyze data";
            this.enableDataHighlight.Name = "enableDataHighlight";
            this.enableDataHighlight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.enableDataHighlight_Click);
            // 
            // ComputeEntropy
            // 
            this.ComputeEntropy.Label = "Compute Entropy";
            this.ComputeEntropy.Name = "ComputeEntropy";
            this.ComputeEntropy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ComputeEntropy_Click);
            // 
            // RegularityMap
            // 
            this.RegularityMap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RegularityMap.Image = global::ExceLintUI.Properties.Resources.ELogo;
            this.RegularityMap.Label = "Reveal Structure";
            this.RegularityMap.Name = "RegularityMap";
            this.RegularityMap.ShowImage = true;
            this.RegularityMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RegularityMap_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::ExceLintUI.Properties.Resources.graph;
            this.button1.Label = "Start Fix";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CheckCellGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VectorForCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RegularityMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox enableDataHighlight;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox analyzeFormulas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ComputeEntropy;
    }

    partial class ThisRibbonCollection
    {
        internal ExceLintRibbon ExceLintRibbon
        {
            get { return this.GetRibbon<ExceLintRibbon>(); }
        }
    }
}
