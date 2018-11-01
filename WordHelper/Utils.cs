using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    internal class Utils {
        internal static bool IsEmptyLine(Word.Paragraph paragraph)
        {
            var isEmpty = true;
            var characters = paragraph.Range.Characters;
            foreach (Word.Range character in characters) {
                if (character.Text == "\r" || character.Text == "\n") {
                    continue;
                }
                isEmpty = false;
            }
            return isEmpty;
        }
    }
}
