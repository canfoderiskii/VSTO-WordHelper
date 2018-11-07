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
            this.WordHelperTab = this.Factory.CreateRibbonTab();
            this.RibbonEditGroup = this.Factory.CreateRibbonGroup();
            this.RibbonVariableGroup = this.Factory.CreateRibbonGroup();
            this.RibbonDevelGroup = this.Factory.CreateRibbonGroup();
            this.RibbonFindReplaceGroup = this.Factory.CreateRibbonGroup();
            this.RibbonFindSelector = this.Factory.CreateRibbonComboBox();
            this.RibbonReplaceSelector = this.Factory.CreateRibbonComboBox();
            this.box1 = this.Factory.CreateRibbonBox();
            this.RibbonFindWildCard = this.Factory.CreateRibbonCheckBox();
            this.RibbonFindMatchCase = this.Factory.CreateRibbonCheckBox();
            this.RibbonFindMatchWholeWord = this.Factory.CreateRibbonCheckBox();
            this.RibbonFindRegex = this.Factory.CreateRibbonCheckBox();
            this.RibbonMenu = this.Factory.CreateRibbonMenu();
            this.RibbonMenuAbout = this.Factory.CreateRibbonButton();
            this.RibbonEditTrimTrailing = this.Factory.CreateRibbonButton();
            this.RibbonEditTrimEmptyLines = this.Factory.CreateRibbonButton();
            this.RibbonEditMergeParagraph = this.Factory.CreateRibbonButton();
            this.RibbonEditConvertLineBreak = this.Factory.CreateRibbonButton();
            this.RibbonVariablePaneToggle = this.Factory.CreateRibbonToggleButton();
            this.RibbonVariableImport = this.Factory.CreateRibbonButton();
            this.RibbonDocVarGenerator = this.Factory.CreateRibbonButton();
            this.RibbonTest = this.Factory.CreateRibbonButton();
            this.RibbonTestDisplayCharCode = this.Factory.CreateRibbonButton();
            this.RibbonFind = this.Factory.CreateRibbonButton();
            this.RibbonReplace = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.WordHelperTab.SuspendLayout();
            this.RibbonEditGroup.SuspendLayout();
            this.RibbonVariableGroup.SuspendLayout();
            this.RibbonDevelGroup.SuspendLayout();
            this.RibbonFindReplaceGroup.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // WordHelperTab
            // 
            this.WordHelperTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.WordHelperTab.Groups.Add(this.RibbonEditGroup);
            this.WordHelperTab.Groups.Add(this.RibbonVariableGroup);
            this.WordHelperTab.Groups.Add(this.RibbonDevelGroup);
            this.WordHelperTab.Groups.Add(this.RibbonFindReplaceGroup);
            this.WordHelperTab.Label = "Word 辅助器";
            this.WordHelperTab.Name = "WordHelperTab";
            // 
            // RibbonEditGroup
            // 
            this.RibbonEditGroup.Items.Add(this.RibbonEditTrimTrailing);
            this.RibbonEditGroup.Items.Add(this.RibbonEditTrimEmptyLines);
            this.RibbonEditGroup.Items.Add(this.RibbonEditMergeParagraph);
            this.RibbonEditGroup.Items.Add(this.RibbonEditConvertLineBreak);
            this.RibbonEditGroup.Label = "文本编辑";
            this.RibbonEditGroup.Name = "RibbonEditGroup";
            // 
            // RibbonVariableGroup
            // 
            this.RibbonVariableGroup.Items.Add(this.RibbonVariablePaneToggle);
            this.RibbonVariableGroup.Items.Add(this.RibbonVariableImport);
            this.RibbonVariableGroup.Label = "文档变量";
            this.RibbonVariableGroup.Name = "RibbonVariableGroup";
            // 
            // RibbonDevelGroup
            // 
            this.RibbonDevelGroup.Items.Add(this.RibbonDocVarGenerator);
            this.RibbonDevelGroup.Items.Add(this.RibbonTest);
            this.RibbonDevelGroup.Items.Add(this.RibbonTestDisplayCharCode);
            this.RibbonDevelGroup.Label = "开发调试";
            this.RibbonDevelGroup.Name = "RibbonDevelGroup";
            // 
            // RibbonFindReplaceGroup
            // 
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonFindSelector);
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonReplaceSelector);
            this.RibbonFindReplaceGroup.Items.Add(this.box1);
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonFindWildCard);
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonFindMatchCase);
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonFindMatchWholeWord);
            this.RibbonFindReplaceGroup.Items.Add(this.RibbonFindRegex);
            this.RibbonFindReplaceGroup.Label = "快速查找替换";
            this.RibbonFindReplaceGroup.Name = "RibbonFindReplaceGroup";
            // 
            // RibbonFindSelector
            // 
            this.RibbonFindSelector.Label = "查找";
            this.RibbonFindSelector.MaxLength = 30;
            this.RibbonFindSelector.Name = "RibbonFindSelector";
            this.RibbonFindSelector.Text = null;
            this.RibbonFindSelector.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonFindSelector_ItemsLoading);
            // 
            // RibbonReplaceSelector
            // 
            this.RibbonReplaceSelector.Label = "修改";
            this.RibbonReplaceSelector.Name = "RibbonReplaceSelector";
            this.RibbonReplaceSelector.Text = null;
            this.RibbonReplaceSelector.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonReplaceSelector_ItemsLoading);
            // 
            // box1
            // 
            this.box1.Items.Add(this.RibbonFind);
            this.box1.Items.Add(this.RibbonReplace);
            this.box1.Name = "box1";
            // 
            // RibbonFindWildCard
            // 
            this.RibbonFindWildCard.Label = "通配符";
            this.RibbonFindWildCard.Name = "RibbonFindWildCard";
            this.RibbonFindWildCard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonFindWildCard_Click);
            // 
            // RibbonFindMatchCase
            // 
            this.RibbonFindMatchCase.Label = "大小写";
            this.RibbonFindMatchCase.Name = "RibbonFindMatchCase";
            // 
            // RibbonFindMatchWholeWord
            // 
            this.RibbonFindMatchWholeWord.Label = "全字匹配";
            this.RibbonFindMatchWholeWord.Name = "RibbonFindMatchWholeWord";
            // 
            // RibbonFindRegex
            // 
            this.RibbonFindRegex.Label = "正则表达式";
            this.RibbonFindRegex.Name = "RibbonFindRegex";
            this.RibbonFindRegex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonFindRegex_Click);
            // 
            // RibbonMenu
            // 
            this.RibbonMenu.Items.Add(this.RibbonMenuAbout);
            this.RibbonMenu.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RibbonMenu.Label = "Word辅助器";
            this.RibbonMenu.Name = "RibbonMenu";
            this.RibbonMenu.OfficeImageId = "CoverPageInsertGallery";
            this.RibbonMenu.ShowImage = true;
            this.RibbonMenu.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonWordHelperMenu_ItemsLoading);
            // 
            // RibbonMenuAbout
            // 
            this.RibbonMenuAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RibbonMenuAbout.Description = "插件信息";
            this.RibbonMenuAbout.Label = "关于";
            this.RibbonMenuAbout.Name = "RibbonMenuAbout";
            this.RibbonMenuAbout.OfficeImageId = "About";
            this.RibbonMenuAbout.ShowImage = true;
            this.RibbonMenuAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonMenuAbout_Click);
            // 
            // RibbonEditTrimTrailing
            // 
            this.RibbonEditTrimTrailing.Label = "清除行尾空白";
            this.RibbonEditTrimTrailing.Name = "RibbonEditTrimTrailing";
            this.RibbonEditTrimTrailing.OfficeImageId = "FormFieldClear";
            this.RibbonEditTrimTrailing.ShowImage = true;
            this.RibbonEditTrimTrailing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditTrimRightButton_Click);
            // 
            // RibbonEditTrimEmptyLines
            // 
            this.RibbonEditTrimEmptyLines.Label = "清除空行";
            this.RibbonEditTrimEmptyLines.Name = "RibbonEditTrimEmptyLines";
            this.RibbonEditTrimEmptyLines.OfficeImageId = "FormFieldClear";
            this.RibbonEditTrimEmptyLines.ShowImage = true;
            this.RibbonEditTrimEmptyLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditTrimEmptyLines_Click);
            // 
            // RibbonEditMergeParagraph
            // 
            this.RibbonEditMergeParagraph.Label = "合并段落";
            this.RibbonEditMergeParagraph.Name = "RibbonEditMergeParagraph";
            this.RibbonEditMergeParagraph.OfficeImageId = "MasterDocumentMergeSubdocuments";
            this.RibbonEditMergeParagraph.ShowImage = true;
            this.RibbonEditMergeParagraph.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditMergeParagraph_Click);
            // 
            // RibbonEditConvertLineBreak
            // 
            this.RibbonEditConvertLineBreak.Label = "转换软回车";
            this.RibbonEditConvertLineBreak.Name = "RibbonEditConvertLineBreak";
            this.RibbonEditConvertLineBreak.OfficeImageId = "MessageNext";
            this.RibbonEditConvertLineBreak.ShowImage = true;
            this.RibbonEditConvertLineBreak.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonEditConvertLineBreak_Click);
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
            // RibbonTestDisplayCharCode
            // 
            this.RibbonTestDisplayCharCode.Label = "显示字符编码";
            this.RibbonTestDisplayCharCode.Name = "RibbonTestDisplayCharCode";
            this.RibbonTestDisplayCharCode.OfficeImageId = "FontDialog";
            this.RibbonTestDisplayCharCode.ShowImage = true;
            this.RibbonTestDisplayCharCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonTestDisplayCharCode_Click);
            // 
            // RibbonFind
            // 
            this.RibbonFind.Label = "查找！";
            this.RibbonFind.Name = "RibbonFind";
            this.RibbonFind.OfficeImageId = "NavigationPaneFind";
            this.RibbonFind.ScreenTip = "RibbonReplace";
            this.RibbonFind.ShowImage = true;
            // 
            // RibbonReplace
            // 
            this.RibbonReplace.Label = "替换！";
            this.RibbonReplace.Name = "RibbonReplace";
            this.RibbonReplace.OfficeImageId = "ReplaceDialog";
            this.RibbonReplace.ShowImage = true;
            this.RibbonReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbonReplace_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            // 
            // Ribbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.RibbonMenu);
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.WordHelperTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.WordHelperTab.ResumeLayout(false);
            this.WordHelperTab.PerformLayout();
            this.RibbonEditGroup.ResumeLayout(false);
            this.RibbonEditGroup.PerformLayout();
            this.RibbonVariableGroup.ResumeLayout(false);
            this.RibbonVariableGroup.PerformLayout();
            this.RibbonDevelGroup.ResumeLayout(false);
            this.RibbonDevelGroup.PerformLayout();
            this.RibbonFindReplaceGroup.ResumeLayout(false);
            this.RibbonFindReplaceGroup.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab WordHelperTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonEditGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonEditTrimTrailing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonEditTrimEmptyLines;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonEditMergeParagraph;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonVariableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton RibbonVariablePaneToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonVariableImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonDevelGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonDocVarGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonEditConvertLineBreak;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox RibbonFindSelector;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonTestDisplayCharCode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox RibbonReplaceSelector;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup RibbonFindReplaceGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox RibbonFindMatchCase;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox RibbonFindMatchWholeWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox RibbonFindWildCard;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox RibbonFindRegex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonFind;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu RibbonMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbonMenuAbout;
    }

    partial class ThisRibbonCollection {
        internal Ribbon Ribbon {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
