using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordHelper {
    internal class Utils {
        internal static int GetPageNumber(Word.Range range)
        {
            return (int)range.Information[Word.WdInformation.wdActiveEndPageNumber];
        }
    }
}
