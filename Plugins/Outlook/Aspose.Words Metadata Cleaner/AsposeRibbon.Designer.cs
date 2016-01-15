namespace Aspose.Words_Metadata_Cleaner
{
    partial class AsposeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AsposeRibbon()
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
            this.AsposeTab = this.Factory.CreateRibbonTab();
            this.AsposeWords = this.Factory.CreateRibbonGroup();
            this.CB_EnableAsposeWordsMetadataCleaner = this.Factory.CreateRibbonCheckBox();
            this.AsposeTab.SuspendLayout();
            this.AsposeWords.SuspendLayout();
            // 
            // AsposeTab
            // 
            this.AsposeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.AsposeTab.Groups.Add(this.AsposeWords);
            this.AsposeTab.Label = "Aspose";
            this.AsposeTab.Name = "AsposeTab";
            // 
            // AsposeWords
            // 
            this.AsposeWords.Items.Add(this.CB_EnableAsposeWordsMetadataCleaner);
            this.AsposeWords.Label = "Aspose.Words Metadata Cleaner";
            this.AsposeWords.Name = "AsposeWords";
            // 
            // CB_EnableAsposeWordsMetadataCleaner
            // 
            this.CB_EnableAsposeWordsMetadataCleaner.Label = "Enable Aspose.Words Metadata Cleaner";
            this.CB_EnableAsposeWordsMetadataCleaner.Name = "CB_EnableAsposeWordsMetadataCleaner";
            this.CB_EnableAsposeWordsMetadataCleaner.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CB_EnableAsposeWordsMetadataCleaner_Click);
            // 
            // AsposeRibbon
            // 
            this.Name = "AsposeRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.AsposeTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AsposeRibbon_Load);
            this.AsposeTab.ResumeLayout(false);
            this.AsposeTab.PerformLayout();
            this.AsposeWords.ResumeLayout(false);
            this.AsposeWords.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AsposeTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AsposeWords;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox CB_EnableAsposeWordsMetadataCleaner;
    }

    partial class ThisRibbonCollection
    {
        internal AsposeRibbon AsposeRibbon
        {
            get { return this.GetRibbon<AsposeRibbon>(); }
        }
    }
}
