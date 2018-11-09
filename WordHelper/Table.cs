using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WordHelper {
    internal static class Table {
        /// <summary>
        /// 检查单元格是否是行合并过的？
        /// </summary>
        internal static bool IsCellRowMerged(Word.Cell cell)
        {
            bool isMerged;
            // 若尝试获取该单元格的下一行单元格失败，那么说明该单元格是行合并过的
            try {
                var cellNextRow = cell.Tables[1].Cell(cell.RowIndex + 1, cell.ColumnIndex);
                // 获取成功，没合并过
                isMerged = false;
            } catch (COMException) { // 获取失败，说明合并过
                isMerged = true;
            }
            return isMerged;
        }
        /// <summary>
        /// 获取单元格真实占据行数量（对于合并单元格这个值大于1）
        /// </summary>
        internal static int GetCellRowSpan(Word.Table table, Word.Cell cell)
        {
            var lastRowIndex = table.Rows.Count;
            var rowCount = 0;
            var cellCount = 0;
            for (var rowIndex = cell.RowIndex; rowIndex <= lastRowIndex; rowIndex++) {
                try {
                    var refCell = table.Cell(rowIndex, cell.ColumnIndex);
                    // 坐标正常流程才能往下
                    if (cellCount < 1) { // 第一次获取成功代表自身
                        rowCount++;
                        cellCount++;
                    } else { // 第二次合并成功代表下一个有效单元格，结束
                        break;
                    }
                } catch (COMException) { // 获取失败，说明合并过
                    rowCount++;
                }
            }
            return rowCount;
        }
        /// <summary>
        /// 获取单元格真实占据列数量（对于合并单元格这个值大于1）
        /// </summary>
        internal static int GetCellColumnSpan(Word.Table table, Word.Cell cell)
        {
            var lastColIndex = table.Columns.Count;
            var colCount = 0;
            var cellCount = 0;
            for (var colIndex = cell.ColumnIndex; colIndex <= lastColIndex; colIndex++) {
                try {
                    var refCell = table.Cell(cell.RowIndex, colIndex);
                    // 坐标正常流程才能往下
                    if (cellCount < 1) { // 第一次获取成功代表自身
                        colCount++;
                        cellCount++;
                    } else { // 第二次合并成功代表下一个有效单元格，结束
                        break;
                    }
                } catch (COMException) { // 获取失败，说明合并过
                    colCount++;
                }
            }
            return colCount;
        }
        /// <summary>
        /// 快速按行拆分单元格为基础单元格
        /// </summary>
        internal static void SplitCellsVertical(Word.Selection selection)
        {
            foreach (Word.Cell cell in selection.Cells) {
                var rowSpan = GetCellRowSpan(selection.Tables[1], cell);
                cell.Split(NumRows: rowSpan);
            }
        }
        /// <summary>
        /// 快速按列拆分单元格为基础单元格
        /// </summary>
        internal static void SplitCellsHorizontal(Word.Selection selection)
        {
            foreach (Word.Cell cell in selection.Cells) {
                var span = GetCellColumnSpan(selection.Tables[1], cell);
                cell.Split(NumColumns: span);
            }
        }
        /// <summary>
        /// 获取当前单元格正下方的单元格
        /// </summary>
        internal static Word.Cell GetNextRowCell(Word.Cell cell)
        {
            return null;
        }

        /// <summary>
        /// 检查单元格是否行合并过且横跨多个页面
        /// </summary>
        internal static bool IsCellMultiPagesMerged(Word.Table table, Word.Cell cell)
        {
            // 判断是否合并过？
            var cellRowSpan = GetCellRowSpan(table, cell);
            if (cellRowSpan <= 1) {
                return false;
            }
            var isMultiPagesMerged = false;
            var cellPageNum = GetCellPageNum(cell);
            // 拆分，为了解析方便
            cell.Split(NumRows: cellRowSpan);
            // 查找每个基本单元格的页码
            for (var curRowIndex = cell.RowIndex + 1; curRowIndex < cell.RowIndex + cellRowSpan; curRowIndex++) {
                var curCell = table.Cell(curRowIndex, cell.ColumnIndex);
                var curCellPageNum = GetCellPageNum(curCell);
                if (curCellPageNum > cellPageNum) { // 找到某个基本单元格在另外一页
                    isMultiPagesMerged = true;
                }
                // 重新合并
                cell.Merge(curCell);
            }

            return isMultiPagesMerged;
        }
        /// <summary>
        /// 获取单元格的起始页码
        /// </summary>
        internal static int GetCellPageNum(Word.Cell cell)
        {
            return (int)cell.Range.Information[Word.WdInformation.wdActiveEndPageNumber];
        }
        internal static void SplitMultiPagesCell(Word.Selection selection)
        {
            // 选中表格？
            if (selection.Tables.Count < 1) {
                return;
            }
            // 选中单个单元格？
            if (selection.Cells.Count > 1) {
                MessageBox.Show("只能选择一个单元格！");
                return;
            }
            // 获取选择的单元格相关信息
            var cell = selection.Cells[1];
            var table = selection.Tables[1]; // 貌似没法通过 Cell 对象获取到关联的 Table
            var cellRowSpan = GetCellRowSpan(table, cell);
            // 不是行合并的单元格？
            if (cellRowSpan <= 1) {
                return;
            }
            // 拆分为基础单元格，为了解析方便
            cell.Split(NumRows: cellRowSpan);
            var cellPageNum = GetCellPageNum(cell);
            var lastRowIndex = cell.RowIndex + cellRowSpan - 1;
            // 直接查看最后一个单元格，快速判定页码范围
            var lastCell = table.Cell(lastRowIndex, cell.ColumnIndex);
            var lastCellPageNum = GetCellPageNum(lastCell);
            var pageSpan = lastCellPageNum - cellPageNum + 1;
            // 如果不跨多页，结束操作
            if (pageSpan < 2) {
                cell.Merge(table.Cell(lastRowIndex, cell.ColumnIndex));
                MessageBox.Show("单元格不跨页！");
                return;
            }
            // 以下肯定跨页
            var splitBaseCells = new List<Word.Cell> { cell };
            var curBaseCell = splitBaseCells[splitBaseCells.Count - 1];
            var curBaseCellPageNum = GetCellPageNum(curBaseCell);
            // 查找每个基本单元格的页码
            for (var curRowIndex = cell.RowIndex + 1; curRowIndex < cell.RowIndex + cellRowSpan; curRowIndex++) {
                var curCell = table.Cell(curRowIndex, curBaseCell.ColumnIndex);
                var curCellPageNum = GetCellPageNum(curCell);
                if (curCellPageNum > curBaseCellPageNum) { // 找到某个基本单元格在另外一页
                    splitBaseCells.Add(curCell); // 该页第一个基本单元格为新的基准单元格
                    // 若已经与最后一格单元格的页码相同，提前结束
                    if (curCellPageNum == lastCellPageNum) {
                        break;
                    }
                    curBaseCell = curCell;
                    curBaseCellPageNum = GetCellPageNum(curCell);
                }
            }
            // 跨页，准备合并、复制内容
            //cell.Range.Copy();
            for (var index = 0; index < splitBaseCells.Count; index++) {
                var baseCell = splitBaseCells[index];
                var rowIndex = 0;

                // 最后一个需要合并的数量为剩余所有
                if (index == splitBaseCells.Count - 1) {
                    rowIndex = lastRowIndex;
                } else { // 其他合并数量为两次基准之间的数量
                    var nextBaseCell = splitBaseCells[index + 1];
                    rowIndex = nextBaseCell.RowIndex - 1;
                }
                baseCell.Merge(table.Cell(rowIndex, cell.ColumnIndex));
                // 复制第一个基准单元格内容到
                if (index > 0) {
                    baseCell.Range.Text = cell.Range.Text;
                    //baseCell.Range.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
                }
            }
        }
    }
}
