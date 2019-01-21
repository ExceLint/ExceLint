﻿namespace ExceLintUI
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
            this.StartOverButton = this.Factory.CreateRibbonButton();
            this.RegularityMap = this.Factory.CreateRibbonButton();
            this.VectorForCell = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ClearEverything = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.button1 = this.Factory.CreateRibbonButton();
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
            this.CheckCellGroup.Items.Add(this.RegularityMap);
            this.CheckCellGroup.Items.Add(this.VectorForCell);
            this.CheckCellGroup.Items.Add(this.separator1);
            this.CheckCellGroup.Items.Add(this.ClearEverything);
            this.CheckCellGroup.Items.Add(this.separator3);
            this.CheckCellGroup.Name = "CheckCellGroup";
            // 
            // box1
            // 
            this.box1.Items.Add(this.AnalyzeButton);
            this.box1.Items.Add(this.MarkAsOKButton);
            this.box1.Items.Add(this.StartOverButton);
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
            this.StartOverButton.Visible = false;
            this.StartOverButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartOverButton_Click);
            // 
            // RegularityMap
            // 
            this.RegularityMap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RegularityMap.Image = global::ExceLintUI.Properties.Resources.graph;
            this.RegularityMap.Label = "Show Global View";
            this.RegularityMap.Name = "RegularityMap";
            this.RegularityMap.ShowImage = true;
            this.RegularityMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RegularityMap_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ClearEverything
            // 
            this.ClearEverything.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ClearEverything.Image = global::ExceLintUI.Properties.Resources.clear_small;
            this.ClearEverything.Label = "Clear Everything";
            this.ClearEverything.Name = "ClearEverything";
            this.ClearEverything.ShowImage = true;
            this.ClearEverything.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClearEverything_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VectorForCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClearEverything;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RegularityMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection
    {
        internal ExceLintRibbon ExceLintRibbon
        {
            get { return this.GetRibbon<ExceLintRibbon>(); }
        }
    }
}
