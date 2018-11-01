using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    internal class Edit {
        internal Edit()
        {
        }
        /// <summary>
        /// 清除选中部分中每个段落结尾的空格/Tab
        /// </summary>
        internal void TrimTrailing(Word.Selection selection)
        {
            var paragraphs = selection.Paragraphs;

            foreach (Word.Paragraph paragraph in paragraphs) {
                var characters = paragraph.Range.Characters;
                var trimStartIndex = -1; // 指向第一个可删除字符
                // 反向查找连续的可删除字符，直接跳过末尾的换行符开始
                for (var i = characters.Count - 1; i > 0; i--) {
                    var character = characters[i];
                    switch (character.Text) {
                    // 先检查空格，因使用`Alt+鼠标`框选模式时，选择面会超过段落末尾。超过末尾的虚拟空格内容也是“ ”。
                    case " ":
                    case "\t":
                        trimStartIndex = i;
                        continue;
                    }
                    // 不处理其他符号
                    break;
                }
                // 若找到了可删除字符索引，则可从它开始删除
                if (trimStartIndex > 0) {
                    characters[trimStartIndex].Delete(Word.WdUnits.wdCharacter, characters.Count - trimStartIndex);
                }
            }
        }
        /// <summary>
        /// 删除选中部分中完全空的段落
        /// </summary>
        internal void TrimEmptyLines(Word.Selection selection)
        {
            foreach (Word.Paragraph paragraph in selection.Paragraphs) {
                if (Utils.IsEmpty(paragraph)) {
                    paragraph.Range.Delete();
                }
            }
        }
    }
}
