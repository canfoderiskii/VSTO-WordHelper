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
                [VariableState.New] = "*",
                [VariableState.Err] = "x",
                [VariableState.Dup] = "+",
            };
        public VariableControl()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 添加所有数据条目行到文档变量中。（跳过空行）
        /// </summary>
        private void AddVariable(DataGridViewRowCollection rows)
        {
            foreach (DataGridViewRow row in rows) {
                var (stateEntry, nameEntry, valueEntry) = this.GetEntry(row);
                // 检查当前是否是空行。对于可编辑的DataGrid，最后一行肯定是空行，需排除
                if (nameEntry.Value == null) {
                    continue;
                }
                // 空值转为空字符串，允许文档出现空的变量值
                var valueString = (string)valueEntry.Value ?? "";
                var nameString = nameEntry.Value.ToString();
                // 跳过已经与内部变量一致的内容，否则Word会报异常
                if (stateEntry.Value.ToString() == _variableStates[VariableState.Sync]) {
                    continue;
                }
                Globals.ThisAddIn.Application.ActiveDocument.Variables.Add(nameString, valueString);
                // 添加成功，修改条目状态标记
                stateEntry.Value = _variableStates[VariableState.Sync];
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private (string state, string name, string val) GetEntryItem(DataGridViewRow row)
        {
            var cell = row.Cells[(int)VariableEntryItem.State];
            var s = (string)(cell.Value ?? _variableStates[VariableState.Err]);
            cell = row.Cells[(int)VariableEntryItem.VarName];
            var n = (string)(cell.Value ?? "");
            cell = row.Cells[(int)VariableEntryItem.VarValue];
            var v = (string)(cell.Value ?? "");

            return (s, n, v);
        }
        private (DataGridViewCell state, DataGridViewCell name, DataGridViewCell val) GetEntry(DataGridViewRow row)
        {
            var cells = row.Cells;
            return (cells[(int)VariableEntryItem.State], cells[(int)VariableEntryItem.VarName], cells[(int)VariableEntryItem.VarValue]);
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
        /// 删除条目和对应的内部变量
        /// </summary>
        private void DelEntryVariable(DataGridViewRow row)
        {
            // 用于填写新数据的新行不能被删除，会引发异常
            if (row.IsNewRow) {
                return;
            }
            // 当前变量已在内部变量中？那么先删除内部变量
            foreach (Word.Variable v in Globals.ThisAddIn.Application.ActiveDocument.Variables) {
                var (_, name, _) = GetEntryItem(row);
                if (v.Name == name) {
                    v.Delete();
                    break;
                }
            }
            VariableDataGrid.Rows.Remove(row); // 从列表中删除指定行
        }
        /// <summary>
        /// 删除符合指定状态的所有条目(不删变量)
        /// </summary>
        internal void DelEntry(VariableState state)
        {
            var rowDelList = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in VariableDataGrid.Rows) {
                if (row.IsNewRow) {
                    continue;
                }
                var (stateString, _, _) = this.GetEntryItem(row);
                // 仅仅添加到删除列表，若直接删除，会造成Rows索引变化，出现不能完全删除的现象
                if (stateString == _variableStates[state]) {
                    rowDelList.Add(row);
                }
            }
            // 批量删除
            foreach (var row in rowDelList) {
                VariableDataGrid.Rows.Remove(row);
            }
        }
        /// <summary>
        /// 使用当前文档内部变量重新生成界面的变量列表内容
        /// </summary>
        internal void ReloadEntry()
        {
            VariableDataGrid.Rows.Clear();
            foreach (Word.Variable v in Globals.ThisAddIn.Application.ActiveDocument.Variables) {
                this.AddEntry(VariableState.Sync, v.Name, v.Value);
            }
        }
        /// <summary>
        /// 让变量列表界面与文档内部变量信息重新同步
        /// 不修改编辑部分
        /// </summary>
        internal void SyncEntry()
        {
            this.DelEntry(VariableState.Sync);
            foreach (Word.Variable v in Globals.ThisAddIn.Application.ActiveDocument.Variables) {
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
            var entryListValid = true;

            foreach (DataGridViewRow row in VariableDataGrid.Rows) {
                // 跳过空行。如果DataGrid可编辑，最后一行为空行，用于用户填入信息的
                if (VariableDataGrid.AllowUserToAddRows && row.Index >= (VariableDataGrid.Rows.Count - 1)) {
                    continue;
                }
                // 获取条目信息
                var (stateCell, nameCell, valCell) = GetEntry(row);
                // 检测是否存在输入不完整的行：仅输入了值，没有名字。（允许仅输入名字，没有值）
                if (nameCell.Value == null) {
                    stateCell.Value = _variableStates[VariableState.Err];
                    entryListValid = false;
                    continue; // 不执行后面代码，因为名字不是字符串
                }

                var varName = nameCell.Value.ToString();
                var varValue = (string)valCell.Value ?? "";
                // 初次出现，添加记录
                if (!statistics.ContainsKey(varName)) {
                    statistics.Add(varName, 0);
                } else { // 再次出现，重复
                    stateCell.Value = _variableStates[VariableState.Dup];
                    entryListValid = false;
                }
                // 命名有效？
                if (!ValidateEntryPattern(varName, varValue)) {
                    stateCell.Value = _variableStates[VariableState.Err];
                    entryListValid = false;
                }
            }
            return entryListValid;
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
                this.DelEntryVariable(row);
            }
        }

        /// <summary>
        /// 文档变量修改确认事件。将当前列表中所有信息更新到文档变量中。
        /// 检查列表是否存在重复信息、无效信息等
        /// </summary>
        private void VariableConfirmButton_Click(object sender, EventArgs e)
        {
            if (!ValidateEntryList()) {
                MessageBox.Show("变量列表有错误，请检查！");
                return;
            }
            this.AddVariable(VariableDataGrid.Rows);
        }
    }
}
