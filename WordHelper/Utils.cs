using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    internal class Utils {
        internal static bool IsEmpty(Word.Paragraph paragraph)
        {
            var isEmpty = true;
            var characters = paragraph.Range.Characters;
            foreach (Word.Range character in characters) {
                var text = character.Text;
                // '\v' 是为了消除存在软回车的段落
                if (text == "\r" || text == "\n" || text == " " || text == "\t" || text == "\v") {
                    continue;
                }
                isEmpty = false;
            }
            return isEmpty;
        }
    }
}
