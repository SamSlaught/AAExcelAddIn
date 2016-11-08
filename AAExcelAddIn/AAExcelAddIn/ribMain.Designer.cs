namespace AAExcelAddIn
{
    partial class ribMain : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ribMain()
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
            this.tabNavigator = this.Factory.CreateRibbonTab();
            this.grpNavigator = this.Factory.CreateRibbonGroup();
            this.btnNavigator = this.Factory.CreateRibbonButton();
            this.tabNavigator.SuspendLayout();
            this.grpNavigator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabNavigator
            // 
            this.tabNavigator.Groups.Add(this.grpNavigator);
            this.tabNavigator.Label = "NAVIGATOR";
            this.tabNavigator.Name = "tabNavigator";
            this.tabNavigator.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // grpNavigator
            // 
            this.grpNavigator.Items.Add(this.btnNavigator);
            this.grpNavigator.Label = "Navigator";
            this.grpNavigator.Name = "grpNavigator";
            // 
            // btnNavigator
            // 
            this.btnNavigator.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNavigator.Label = "Navigator";
            this.btnNavigator.Name = "btnNavigator";
            this.btnNavigator.OfficeImageId = "FindDialog";
            this.btnNavigator.ShowImage = true;
            this.btnNavigator.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavigator_Click);
            // 
            // ribMain
            // 
            this.Name = "ribMain";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabNavigator);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabNavigator.ResumeLayout(false);
            this.tabNavigator.PerformLayout();
            this.grpNavigator.ResumeLayout(false);
            this.grpNavigator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabNavigator;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpNavigator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavigator;
    }

    partial class ThisRibbonCollection
    {
        internal ribMain Ribbon1
        {
            get { return this.GetRibbon<ribMain>(); }
        }
    }
}
