namespace WordHelper {
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            if (disposing && (components != null)) {
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
            this.RibbonDocVarGroup = this.Factory.CreateRibbonGroup();
            this.RibbonVariablePaneToggle = this.Factory.CreateRibbonToggleButton();
            this.RibbonDocVarImport = this.Factory.CreateRibbonButton();
            this.RibbonDevelGroup = this.Factory.CreateRibbonGroup();
            this.RibbonDocVarGenerator = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.RibbonDocVarGroup.SuspendLayout();
            this.RibbonDevelGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.RibbonDocVarGroup);
            this.tab1.Groups.Add(this.RibbonDevelGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // RibbonDocVarGroup
            // 
            this.RibbonDocVarGroup.Items.Add(this.RibbonVariablePaneToggle);
            this.RibbonDocVarGroup.Items.Add(this.RibbonDocVarImport);
            this.RibbonDocVarGroup.Label = "文档变量";
            this.RibbonDocVarGroup.Name = "RibbonDocVarGroup";
            // 
            // RibbonVariablePaneToggle
            // 
            this.RibbonVariablePaneToggle.Label = "显示变量";
            this.RibbonVariablePaneToggle.Name = "RibbonVariablePaneToggle";
            this.RibbonVariablePaneToggle.OfficeImageId = "ViewDraftView";
            this.RibbonVariablePaneToggle.ShowImage = true;
            this.RibbonVariablePaneToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariablePaneToggle_Click);
            // 
            // RibbonDocVarImport
            // 
            this.RibbonDocVarImport.Label = "从文件导入";
            this.RibbonDocVarImport.Name = "RibbonDocVarImport";
            this.RibbonDocVarImport.OfficeImageId = "MailMergeDocument";
            this.RibbonDocVarImport.ShowImage = true;
            this.RibbonDocVarImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariableImport_Click);
            // 
            // RibbonDevelGroup
            // 
            this.RibbonDevelGroup.Items.Add(this.RibbonDocVarGenerator);
            this.RibbonDevelGroup.Label = "开发调试";
            this.RibbonDevelGroup.Name = "RibbonDevelGroup";
            // 
            // RibbonDocVarGenerator
            // 
            this.RibbonDocVarGenerator.Label = "生成变量";
            this.RibbonDocVarGenerator.Name = "RibbonDocVarGenerator";
            this.RibbonDocVarGenerator.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariableGenerator_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.RibbonDocVarGroup.ResumeLayout(false);
            this.RibbonDocVarGroup.PerformLayout();
            this.RibbonDevelGroup.ResumeLayout(false);
            this.RibbonDevelGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonDocVarGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton RibbonVariablePaneToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonDocVarImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonDevelGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonDocVarGenerator;
    }

    partial class ThisRibbonCollection {
        internal Ribbon Ribbon {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
