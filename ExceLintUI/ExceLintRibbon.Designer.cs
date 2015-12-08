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
            this.SensitivityTextBox = this.Factory.CreateRibbonEditBox();
            this.AnalyzeButton = this.Factory.CreateRibbonButton();
            this.MarkAsOKButton = this.Factory.CreateRibbonButton();
            this.FixErrorButton = this.Factory.CreateRibbonButton();
            this.StartOverButton = this.Factory.CreateRibbonButton();
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
            this.CheckCellGroup.Items.Add(this.SensitivityTextBox);
            this.CheckCellGroup.Name = "CheckCellGroup";
            // 
            // box1
            // 
            this.box1.Items.Add(this.AnalyzeButton);
            this.box1.Items.Add(this.MarkAsOKButton);
            this.box1.Items.Add(this.FixErrorButton);
            this.box1.Items.Add(this.StartOverButton);
            this.box1.Name = "box1";
            // 
            // SensitivityTextBox
            // 
            this.SensitivityTextBox.Label = "% to Audit";
            this.SensitivityTextBox.Name = "SensitivityTextBox";
            this.SensitivityTextBox.SizeString = "100.0";
            this.SensitivityTextBox.Text = "5.0";
            // 
            // AnalyzeButton
            // 
            this.AnalyzeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AnalyzeButton.Image = global::ExceLintUI.Properties.Resources.analyze_small;
            this.AnalyzeButton.Label = "Analyze";
            this.AnalyzeButton.Name = "AnalyzeButton";
            this.AnalyzeButton.ShowImage = true;
            // 
            // MarkAsOKButton
            // 
            this.MarkAsOKButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.MarkAsOKButton.Image = global::ExceLintUI.Properties.Resources.mark_as_ok_small;
            this.MarkAsOKButton.Label = "Mark as OK";
            this.MarkAsOKButton.Name = "MarkAsOKButton";
            this.MarkAsOKButton.ShowImage = true;
            // 
            // FixErrorButton
            // 
            this.FixErrorButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FixErrorButton.Image = global::ExceLintUI.Properties.Resources.correct_small;
            this.FixErrorButton.Label = "Fix Error";
            this.FixErrorButton.Name = "FixErrorButton";
            this.FixErrorButton.ShowImage = true;
            // 
            // StartOverButton
            // 
            this.StartOverButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StartOverButton.Image = global::ExceLintUI.Properties.Resources.clear_small;
            this.StartOverButton.Label = "Start Over";
            this.StartOverButton.Name = "StartOverButton";
            this.StartOverButton.ShowImage = true;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SensitivityTextBox;
    }

    partial class ThisRibbonCollection
    {
        internal ExceLintRibbon ExceLintRibbon
        {
            get { return this.GetRibbon<ExceLintRibbon>(); }
        }
    }
}
