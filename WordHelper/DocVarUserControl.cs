using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    public partial class DocVarUserControl : UserControl {
        public DocVarUserControl()
        {
            InitializeComponent();
        }

        internal void DocVarReloadDataGrid()
        {
            DocVarDataGrid.Rows.Clear();

            var document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Microsoft.Office.Interop.Word.Variable v in document.Variables) {
                DocVarDataGrid.Rows.Add(new string[] { v.Name, v.Value });
            }
        }

        private void DocVarReloadButton_Click(object sender, EventArgs e)
        {
            this.DocVarReloadDataGrid();
        }

        /// <summary>
        /// 右键菜单删除点击动作，删除选中的所有行。如果当前有未保存的文档变量修改状态，提示用户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DocVarContextDelete_Click(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            // TODO: 检查是否存在修改状态
            foreach (DataGridViewRow row in DocVarDataGrid.SelectedRows) {
                foreach (Word.Variable v in document.Variables) {
                    if (v.Name == (string)row.Cells[0].Value) {
                        v.Delete();
                        break;
                    }
                }
                DocVarDataGrid.Rows.Remove(row);
            }
        }
    }
}
