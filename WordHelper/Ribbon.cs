using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;

namespace WordHelper {
    public partial class Ribbon {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonDocVarPaneToggle.Checked = Globals.ThisAddIn.DocVarPane.Visible;
        }

        private void RibbonDocVarPaneToggle_Click(object sender, RibbonControlEventArgs e)
        {
            var pane = Globals.ThisAddIn.DocVarPane;
            pane.Visible = !pane.Visible;
        }

        /// <summary>
        /// 导入外部文档内容按钮触发。弹出对话框选择文档，将Excel第一页工作表内容导入到当前显示列表中，无误后人工确认
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibbonDocVarImport_Click(object sender, RibbonControlEventArgs e)
        {
            var dlg = new OpenFileDialog {
                DefaultExt = ".xlsx",
                Filter = "Excel文件|*.xlsx"
            };
            if (dlg.ShowDialog() == true) {
                var filename = dlg.FileName;
                // 尝试打开 Excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filename);
                Excel.Worksheet sheet = workbook.Sheets[1];
                Excel.Range usedRange = sheet.UsedRange;
                do {
                    // 检查基本格式
                    if (usedRange.Columns.Count != 2) {
                        MessageBox.Show("Excel工作表内容格式不正确：应该只有两列！");
                        break;
                    }

                    // 文档内部变量有可能没有显示到变量列表控件中，需要先同步一次


                    // 遍历所有行，第1列为文档变量名，第2列为文档变量值
                    foreach (Excel.Range rows in usedRange.Rows) {
                        var colCell1 = (Excel.Range)rows.Cells.Item[1, 1];
                        var colCell2 = (Excel.Range)rows.Cells.Item[1, 2];
                        var docVarName = colCell1.Value.ToString();
                        var docVarVal = colCell2.Value.ToString();
                        var control = (DocVarUserControl)Globals.ThisAddIn.DocVarPane.Control;
                        control.AddDataGridItem(DocVarDataGridState.New, docVarName, docVarVal);
                    }
                } while (false);
                // 处理收尾
                excelApp.Quit();
            }
        }

        #region 内部开发调试

        private static uint _count = 0;

        private void RibbonDocVarGenerator_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Variables.Add("TESTVAR" + _count, "TESTVALUE" + _count);
            _count++;
        }

        #endregion
    }
}
