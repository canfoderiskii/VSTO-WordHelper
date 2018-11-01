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

        internal void TrimTrailing(Word.Selection selection)
        {
            foreach (Word.Paragraph paragraph in selection.Paragraphs) {
                var characters = paragraph.Range.Characters;
                for (var i = characters.Count; i > 0; i--) {
                    var character = characters[i];
                    // 先检查空格，因使用`Alt+鼠标`框选模式时，选择面会超过段落末尾。超过末尾的虚拟空格内容也是“ ”。
                    if (character.Text == " ") {
                        character.Delete(Word.WdUnits.wdCharacter, 1);
                        continue;
                    }
                    // 跳过换行符
                    if (character.Text == "\r" || character.Text == "\n") {
                        continue;
                    }
                    // 不处理其他符号
                    break;
                }
            }
        }

        internal void TrimEmptyLines(Word.Selection selection)
        {
            foreach (Word.Paragraph paragraph in selection.Paragraphs) {
                if (Utils.IsEmptyLine(paragraph)) {
                    paragraph.Range.Delete();
                }
            }
        }
    }
}
