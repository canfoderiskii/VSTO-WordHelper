using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using OfficeRibbon = Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;

namespace WordHelper {
    public partial class Ribbon {
        private static readonly Dictionary<string, string> _ribbonFindReplacePresets = new Dictionary<string, string>() {
            ["段落标记"] = "^p",
            ["换行标记"] = "^l",
        };

        private void Ribbon_LoadFindReplaceDropDownItems()
        {
            // 不同的 DropDown 必须加入不同的 DropDownItem
            foreach (var dictItem in _ribbonFindReplacePresets) {
                var item0 = this.Factory.CreateRibbonDropDownItem();
                var item1 = this.Factory.CreateRibbonDropDownItem();
                item0.Label = dictItem.Value;
                item0.ScreenTip = dictItem.Key;
                item1.Label = dictItem.Key;
                item1.ScreenTip = dictItem.Key;
                RibbonFindSelector.Items.Add(item0);
                RibbonReplaceSelector.Items.Add(item1);
            }
        }
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonVariablePaneToggle.Checked = Globals.ThisAddIn.VariablePane.Visible;
            Ribbon_LoadFindReplaceDropDownItems();
        }

        private void RibbonVariablePaneToggle_Click(object sender, RibbonControlEventArgs e)
        {
            var pane = Globals.ThisAddIn.VariablePane;
            pane.Visible = !pane.Visible;
        }

        /// <summary>
        /// 导入外部文档内容按钮触发。弹出对话框选择文档，将Excel第一页工作表内容导入到当前显示列表中，无误后人工确认
        /// </summary>
        private void RibbonVariableImport_Click(object sender, RibbonControlEventArgs e)
        {
            var dlg = new OpenFileDialog {
                DefaultExt = ".xlsx",
                Filter = "Excel文件|*.xlsx"
            };
            if (dlg.ShowDialog() != true)
                return;

            var filename = dlg.FileName;
            // 尝试打开 Excel
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(filename);
            var sheet = workbook.Sheets[1];
            var usedRange = sheet.UsedRange;
            do {
                // 检查基本格式
                if (usedRange.Columns.Count != 2) {
                    MessageBox.Show("Excel工作表内容格式不正确：应该只有两列！");
                    break;
                }
                // 文档内部变量有可能没有显示到变量列表控件中，需要先同步一次
                Globals.ThisAddIn.VariableControl.SyncEntry();
                // 遍历所有行，第1列为文档变量名，第2列为文档变量值
                var variableCollection = new Dictionary<string, string>();
                foreach (Excel.Range rows in usedRange.Rows) {
                    var colCell1 = (Excel.Range)rows.Cells.Item[1, 1];
                    var colCell2 = (Excel.Range)rows.Cells.Item[1, 2];
                    var varName = colCell1.Value.ToString();
                    var varVal = colCell2.Value.ToString();

                    if (variableCollection.ContainsKey(varName)) {
                        MessageBox.Show("导入文件中有重复变量定义，请检查！");
                        continue;
                    }
                    variableCollection.Add(varName, varVal);
                }
                foreach (var index in variableCollection) {
                    Globals.ThisAddIn.VariableControl.AddEntry(VariableState.New, index.Key, index.Value);
                }
            } while (false);
            // 处理收尾
            excelApp.Quit();
        }

        #region 内部开发调试
        private static uint _count = 0;

        private void RibbonVariableGenerator_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Variables.Add("TESTVAR" + _count, "TESTVALUE" + _count);
            _count++;
        }
        private void RibbonTest_Click(object sender, RibbonControlEventArgs e)
        {
        }
        #endregion

        #region 文本编辑相关功能
        private void RibbonEditTrimRightButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Edit.TrimTrailing(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }
        private void RibbonEditTrimEmptyLines_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Edit.TrimEmptyLines(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }
        private void RibbonEditMergeParagraph_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Edit.MergeParagraph(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }
        private void RibbonEditConvertLineBreak_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Edit.ConvertLineBreak(Globals.ThisAddIn.Application.ActiveWindow.Selection);
        }
        private void RibbonFindReplace_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            var find = selection.Range.Find;
            var findText = this.RibbonFindSelector.Text;
            var replaceText = this.RibbonReplaceSelector.Text;

            // 若没有可查找的直接退出
            if (findText == "") {
                return;
            }
            // 若查询
            //if (_ribbonFindReplacePresets.ContainsKey(findText)) {

            //}

            find.Execute(FindText: findText, MatchWholeWord: this.RibbonFindMatchWholeWord.Checked, Forward: false, Wrap: Word.WdFindWrap.wdFindStop, Replace: Word.WdReplace.wdReplaceAll, ReplaceWith: replaceText);
        }
        #endregion
    }
}
