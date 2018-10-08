using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    internal enum DocVarDataGridState {
        Sync,
        New,
        Err,
        Dup,
    };

    internal enum DocVarDataGridCol {
        State = 0,
        VarName = 1,
        VarValue = 2,
    };

    public partial class DocVarUserControl : UserControl {
        private readonly Dictionary<DocVarDataGridState, string> _dataGridStates
            = new Dictionary<DocVarDataGridState, string>() {
                [DocVarDataGridState.Sync] = "",
                [DocVarDataGridState.New] = "^",
                [DocVarDataGridState.Err] = "x",
                [DocVarDataGridState.Dup] = "+",
            };
        public DocVarUserControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 添加文档变量名和值到界面的DataGrid中。
        /// “不检查”是否存在重复
        /// </summary>
        internal void AddDataGridItem(DocVarDataGridState dataState, string varName, string varValue)
        {
            string[] item = { _dataGridStates[dataState], varName, varValue };
            DocVarDataGrid.Rows.Add(item);
        }

        internal void ReloadDataGrid()
        {
            DocVarDataGrid.Rows.Clear();
            var document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Microsoft.Office.Interop.Word.Variable v in document.Variables) {
                this.AddDataGridItem(DocVarDataGridState.Sync, v.Name, v.Value);
            }
        }

        /// <summary>
        /// 检查文档变量名是否有效
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varValue"></param>
        /// <returns>true, 有效； false, 无效</returns>
        internal bool ValidateDocVarPattern(string varName, string varValue)
        {
            const string pattern = @"^[1-9]"; // 不能以数字开头
            var regex = new Regex(pattern);
            return regex.Matches(varName).Count <= 0;
        }
        /// <summary>
        /// 检查文档变量展示数据表内容是否符合要求
        /// </summary>
        /// <returns>true, 有效； false, 无效</returns>
        internal bool ValidateDocVarDataGrid()
        {
            var statistics = new Dictionary<string, int>();
            var dataGridValid = true;

            foreach (DataGridViewRow row in DocVarDataGrid.Rows) {
                // 跳过空行。如果DataGrid可编辑，最后一行为空行，用于用户填入信息的
                if (DocVarDataGrid.AllowUserToAddRows && row.Index >= (DocVarDataGrid.Rows.Count - 1)) {
                    continue;
                }

                var stateCell = row.Cells[(int)DocVarDataGridCol.State];
                var varName = row.Cells[(int)DocVarDataGridCol.VarName].Value.ToString();
                var varValue = row.Cells[(int)DocVarDataGridCol.VarValue].Value.ToString();
                // 初次出现，添加记录
                if (!statistics.ContainsKey(varName)) {
                    statistics.Add(varName, 0);
                } else { // 再次出现，重复
                    stateCell.Value = _dataGridStates[DocVarDataGridState.Dup];
                    dataGridValid = false;
                }
                // 命名有效？
                if (!ValidateDocVarPattern(varName, varValue)) {
                    stateCell.Value = _dataGridStates[DocVarDataGridState.Err];
                    dataGridValid = false;
                }
            }
            return dataGridValid;
        }

        // TODO: 添加搜寻接口，得出统计信息

        private void DocVarReloadButton_Click(object sender, EventArgs e)
        {
            this.ReloadDataGrid();
        }

        /// <summary>
        /// 右键菜单删除点击动作，删除选中的所有行。如果当前有未保存的文档变量修改状态，提示用户
        /// </summary>
        private void DocVarContextDelete_Click(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            // TODO: 检查是否存在修改状态
            foreach (DataGridViewRow row in DocVarDataGrid.SelectedRows) {
                foreach (Word.Variable v in document.Variables) {
                    const int i = (int)DocVarDataGridCol.VarName;
                    if (v.Name == (string)row.Cells[i].Value) {
                        v.Delete();
                        break;
                    }
                }
                DocVarDataGrid.Rows.Remove(row);
            }
        }

        /// <summary>
        /// 文档变量修改确认事件。将当前列表中所有信息更新到文档变量中。
        /// 检查列表是否存在重复信息、无效信息等
        /// </summary>
        private void DocVarConfirmButton_Click(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            if (!ValidateDocVarDataGrid()) {
                MessageBox.Show("变量列表有错误，请检查！");
                return;
            }

            foreach (DataGridViewRow row in DocVarDataGrid.Rows) {
                const int nameIdx = (int)DocVarDataGridCol.VarName;
                const int valueIdx = (int)DocVarDataGridCol.VarValue;

                // 跳过已经与内部变量一致的内容，否则Word会报异常
                if (row.Cells[(int)DocVarDataGridCol.State].Value.ToString() == _dataGridStates[DocVarDataGridState.Sync]) {
                    continue;
                    ;
                }

                document.Variables.Add(row.Cells[nameIdx].Value.ToString(), row.Cells[valueIdx].Value);
            }
        }
    }
}
