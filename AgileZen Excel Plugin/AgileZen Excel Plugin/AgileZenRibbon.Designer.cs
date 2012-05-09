namespace AgileZen_Excel_Plugin
{
    partial class AgileZenRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AgileZenRibbon()
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
            this.AgileZenGroup = this.Factory.CreateRibbonGroup();
            this.ProjectsDropDown = this.Factory.CreateRibbonDropDown();
            this.PhaseDropDown = this.Factory.CreateRibbonDropDown();
            this.LogonButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AgileZenGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.AgileZenGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // AgileZenGroup
            // 
            this.AgileZenGroup.Items.Add(this.LogonButton);
            this.AgileZenGroup.Items.Add(this.ProjectsDropDown);
            this.AgileZenGroup.Items.Add(this.PhaseDropDown);
            this.AgileZenGroup.Label = "AgileZen";
            this.AgileZenGroup.Name = "AgileZenGroup";
            // 
            // ProjectsDropDown
            // 
            this.ProjectsDropDown.Label = "Prosjekter";
            this.ProjectsDropDown.Name = "ProjectsDropDown";
            this.ProjectsDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.projectDropDown_SelectionChanged);
            // 
            // PhaseDropDown
            // 
            this.PhaseDropDown.Label = "Fase";
            this.PhaseDropDown.Name = "PhaseDropDown";
            this.PhaseDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PhaseDropDown_SelectionChanged);
            // 
            // LogonButton
            // 
            this.LogonButton.Label = "Logg på";
            this.LogonButton.Name = "LogonButton";
            this.LogonButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogonButton_Click);
            // 
            // AgileZenRibbon
            // 
            this.Name = "AgileZenRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AgileZenGroup.ResumeLayout(false);
            this.AgileZenGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AgileZenGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ProjectsDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown PhaseDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LogonButton;
    }

    partial class ThisRibbonCollection
    {
        internal AgileZenRibbon Ribbon1
        {
            get { return this.GetRibbon<AgileZenRibbon>(); }
        }
    }
}
