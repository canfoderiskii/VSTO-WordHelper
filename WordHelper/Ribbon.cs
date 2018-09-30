using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordHelper {
    public partial class Ribbon {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            RibbonDocVarPaneToggle.Checked = Globals.ThisAddIn.DocVarPane.Visible;
        }

        private void RibbonDocVarPaneToggle_Click(object sender, RibbonControlEventArgs e)
        {
            var pane = Globals.ThisAddIn.DocVarPane;
            pane.Visible = !pane.Visible;
        }

        private void RibbonDocVarImport_Click(object sender, RibbonControlEventArgs e)
        {

        }

        #region 内部开发调试

        private static uint _count = 0;

        private void RibbonDocVarGenerator_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Variables.Add("TESTVAR" + _count, "TESTVALUE" + _count);
            _count++;
        }

        #endregion
    }
}
