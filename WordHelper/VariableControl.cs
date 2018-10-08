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
    /// <summary>
    /// 文档变量列表中每项条目可能的状态
    /// </summary>
    internal enum VariableState {
        Sync,
        New,
        Err,
        Dup,
    };
    /// <summary>
    /// 文档变量列表条目各信息标识
    /// </summary>
    internal enum VariableEntryItem {
        State = 0,
        VarName = 1,
        VarValue = 2,
    };

    public partial class VariableControl : UserControl {
        private readonly Dictionary<VariableState, string> _variableStates
            = new Dictionary<VariableState, string>() {
                [VariableState.Sync] = "",
                [VariableState.New] = "^",
                [VariableState.Err] = "x",
                [VariableState.Dup] = "+",
            };
        public VariableControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 添加文档变量条目到界面的列表中
        /// “不检查”是否存在重复
        /// </summary>
        internal void AddEntry(VariableState varState, string varName, string varValue)
        {
            string[] item = { _variableStates[varState], varName, varValue };
            VariableDataGrid.Rows.Add(item);
        }
        /// <summary>
        /// 使用当前文档内部变量重新生成界面的变量列表内容
        /// </summary>
        internal void ReloadEntry()
        {
            VariableDataGrid.Rows.Clear();
            var document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Microsoft.Office.Interop.Word.Variable v in document.Variables) {
                this.AddEntry(VariableState.Sync, v.Name, v.Value);
            }
        }
        /// <summary>
        /// 检查文档变量名是否有效
        /// </summary>
        /// <returns>true, 有效； false, 无效</returns>
        internal bool ValidateEntryPattern(string varName, string varValue)
        {
            const string pattern = @"^[1-9]"; // 不能以数字开头
            var regex = new Regex(pattern);
            return regex.Matches(varName).Count <= 0;
        }
        /// <summary>
        /// 检查文档变量展示数据表内容是否符合要求
        /// </summary>
        /// <returns>true, 有效； false, 无效</returns>
        internal bool ValidateEntryList()
        {
            var statistics = new Dictionary<string, int>();
            var dataGridValid = true;

            foreach (DataGridViewRow row in VariableDataGrid.Rows) {
                // 跳过空行。如果DataGrid可编辑，最后一行为空行，用于用户填入信息的
                if (VariableDataGrid.AllowUserToAddRows && row.Index >= (VariableDataGrid.Rows.Count - 1)) {
                    continue;
                }

                var stateCell = row.Cells[(int)VariableEntryItem.State];
                var varName = row.Cells[(int)VariableEntryItem.VarName].Value.ToString();
                var varValue = row.Cells[(int)VariableEntryItem.VarValue].Value.ToString();
                // 初次出现，添加记录
                if (!statistics.ContainsKey(varName)) {
                    statistics.Add(varName, 0);
                } else { // 再次出现，重复
                    stateCell.Value = _variableStates[VariableState.Dup];
                    dataGridValid = false;
                }
                // 命名有效？
                if (!ValidateEntryPattern(varName, varValue)) {
                    stateCell.Value = _variableStates[VariableState.Err];
                    dataGridValid = false;
                }
            }
            return dataGridValid;
        }

        // TODO: 添加搜寻接口，得出统计信息

        private void VariableReloadButton_Click(object sender, EventArgs e)
        {
            this.ReloadEntry();
        }

        /// <summary>
        /// 右键菜单删除点击动作，删除选中的所有行。如果当前有未保存的文档变量修改状态，提示用户
        /// </summary>
        private void VariableContextDelete_Click(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            // TODO: 检查是否存在修改状态
            foreach (DataGridViewRow row in VariableDataGrid.SelectedRows) {
                foreach (Word.Variable v in document.Variables) {
                    const int i = (int)VariableEntryItem.VarName;
                    if (v.Name == (string)row.Cells[i].Value) {
                        v.Delete();
                        break;
                    }
                }
                VariableDataGrid.Rows.Remove(row);
            }
        }

        /// <summary>
        /// 文档变量修改确认事件。将当前列表中所有信息更新到文档变量中。
        /// 检查列表是否存在重复信息、无效信息等
        /// </summary>
        private void VariableConfirmButton_Click(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            if (!ValidateEntryList()) {
                MessageBox.Show("变量列表有错误，请检查！");
                return;
            }

            foreach (DataGridViewRow row in VariableDataGrid.Rows) {
                const int nameIdx = (int)VariableEntryItem.VarName;
                const int valueIdx = (int)VariableEntryItem.VarValue;

                // 跳过已经与内部变量一致的内容，否则Word会报异常
                if (row.Cells[(int)VariableEntryItem.State].Value.ToString() == _variableStates[VariableState.Sync]) {
                    continue;
                }
                document.Variables.Add(row.Cells[nameIdx].Value.ToString(), row.Cells[valueIdx].Value);
            }
        }
    }
}
