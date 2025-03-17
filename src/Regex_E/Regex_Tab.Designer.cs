namespace Regex_E
{
    partial class Regex_Tab : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Regex_Tab()
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
            this.groupMain = this.Factory.CreateRibbonGroup();
            this.btnGroupMain = this.Factory.CreateRibbonButtonGroup();
            this.btnOpenDialog = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupMain.SuspendLayout();
            this.btnGroupMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupMain);
            this.tab1.Label = "Regex_E";
            this.tab1.Name = "tab1";
            // 
            // groupMain
            // 
            this.groupMain.Items.Add(this.btnGroupMain);
            this.groupMain.Label = "main";
            this.groupMain.Name = "groupMain";
            // 
            // btnGroupMain
            // 
            this.btnGroupMain.Items.Add(this.btnOpenDialog);
            this.btnGroupMain.Name = "btnGroupMain";
            // 
            // btnOpenDialog
            // 
            this.btnOpenDialog.Label = "more";
            this.btnOpenDialog.Name = "btnOpenDialog";
            this.btnOpenDialog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMain_Click);
            // 
            // Regex_Tab
            // 
            this.Name = "Regex_Tab";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Regex_Tab_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupMain.ResumeLayout(false);
            this.groupMain.PerformLayout();
            this.btnGroupMain.ResumeLayout(false);
            this.btnGroupMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup btnGroupMain;
    }

    partial class ThisRibbonCollection
    {
        internal Regex_Tab Regex_Tab
        {
            get { return this.GetRibbon<Regex_Tab>(); }
        }
    }
}
