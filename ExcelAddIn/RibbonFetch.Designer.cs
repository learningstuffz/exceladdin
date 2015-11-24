namespace ExcelAddIn
{
    partial class RibbonFetch : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonFetch()
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
            this.grpData = this.Factory.CreateRibbonGroup();
            this.btnFetch = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpData.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpData);
            this.tab1.Label = "Fetch-Sample";
            this.tab1.Name = "tab1";
            // 
            // grpData
            // 
            this.grpData.Items.Add(this.btnFetch);
            this.grpData.Label = "LoadData";
            this.grpData.Name = "grpData";
            // 
            // btnFetch
            // 
            this.btnFetch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFetch.Image = global::ExcelAddIn.Properties.Resources.dwnload;
            this.btnFetch.ImageName = "download";
            this.btnFetch.Label = "Fetch Data";
            this.btnFetch.Name = "btnFetch";
            this.btnFetch.ShowImage = true;
            this.btnFetch.SuperTip = "Click to Fetch Data";
            this.btnFetch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFetch_Click);
            // 
            // RibbonFetch
            // 
            this.Name = "RibbonFetch";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonFetch_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpData.ResumeLayout(false);
            this.grpData.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFetch;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonFetch RibbonFetch
        {
            get { return this.GetRibbon<RibbonFetch>(); }
        }
    }
}
