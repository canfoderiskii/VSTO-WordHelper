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
            this.RibbonVariableGroup = this.Factory.CreateRibbonGroup();
            this.RibbonVariablePaneToggle = this.Factory.CreateRibbonToggleButton();
            this.RibbonVariableImport = this.Factory.CreateRibbonButton();
            this.RibbonDevelGroup = this.Factory.CreateRibbonGroup();
            this.RibbonDocVarGenerator = this.Factory.CreateRibbonButton();
            this.RibbonTest = this.Factory.CreateRibbonButton();
            this.RibbonEditGroup = this.Factory.CreateRibbonGroup();
            this.RibbonEditTrimTrailing = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.RibbonVariableGroup.SuspendLayout();
            this.RibbonDevelGroup.SuspendLayout();
            this.RibbonEditGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.RibbonEditGroup);
            this.tab1.Groups.Add(this.RibbonVariableGroup);
            this.tab1.Groups.Add(this.RibbonDevelGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // RibbonVariableGroup
            // 
            this.RibbonVariableGroup.Items.Add(this.RibbonVariablePaneToggle);
            this.RibbonVariableGroup.Items.Add(this.RibbonVariableImport);
            this.RibbonVariableGroup.Label = "文档变量";
            this.RibbonVariableGroup.Name = "RibbonVariableGroup";
            // 
            // RibbonVariablePaneToggle
            // 
            this.RibbonVariablePaneToggle.Label = "显示变量";
            this.RibbonVariablePaneToggle.Name = "RibbonVariablePaneToggle";
            this.RibbonVariablePaneToggle.OfficeImageId = "ViewDraftView";
            this.RibbonVariablePaneToggle.ShowImage = true;
            this.RibbonVariablePaneToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariablePaneToggle_Click);
            // 
            // RibbonVariableImport
            // 
            this.RibbonVariableImport.Label = "从文件导入";
            this.RibbonVariableImport.Name = "RibbonVariableImport";
            this.RibbonVariableImport.OfficeImageId = "MailMergeDocument";
            this.RibbonVariableImport.ShowImage = true;
            this.RibbonVariableImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariableImport_Click);
            // 
            // RibbonDevelGroup
            // 
            this.RibbonDevelGroup.Items.Add(this.RibbonDocVarGenerator);
            this.RibbonDevelGroup.Items.Add(this.RibbonTest);
            this.RibbonDevelGroup.Label = "开发调试";
            this.RibbonDevelGroup.Name = "RibbonDevelGroup";
            // 
            // RibbonDocVarGenerator
            // 
            this.RibbonDocVarGenerator.Label = "生成变量";
            this.RibbonDocVarGenerator.Name = "RibbonDocVarGenerator";
            this.RibbonDocVarGenerator.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonVariableGenerator_Click);
            // 
            // RibbonTest
            // 
            this.RibbonTest.Label = "测试";
            this.RibbonTest.Name = "RibbonTest";
            this.RibbonTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonTest_Click);
            // 
            // RibbonEditGroup
            // 
            this.RibbonEditGroup.Items.Add(this.RibbonEditTrimTrailing);
            this.RibbonEditGroup.Label = "文本编辑";
            this.RibbonEditGroup.Name = "RibbonEditGroup";
            // 
            // RibbonEditTrimTrailing
            // 
            this.RibbonEditTrimTrailing.Label = "清除行尾空白";
            this.RibbonEditTrimTrailing.Name = "RibbonEditTrimTrailing";
            this.RibbonEditTrimTrailing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditTrimRightButton_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.RibbonVariableGroup.ResumeLayout(false);
            this.RibbonVariableGroup.PerformLayout();
            this.RibbonDevelGroup.ResumeLayout(false);
            this.RibbonDevelGroup.PerformLayout();
            this.RibbonEditGroup.ResumeLayout(false);
            this.RibbonEditGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonVariableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton RibbonVariablePaneToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonVariableImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonDevelGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonDocVarGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonEditGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonEditTrimTrailing;
    }

    partial class ThisRibbonCollection {
        internal Ribbon Ribbon {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
