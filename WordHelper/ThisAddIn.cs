﻿using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordHelper {
    public partial class ThisAddIn {
        internal VariableControl VariableControl { get; set; } = new VariableControl();
        internal AboutBox AboutBox { get; set; } = new AboutBox();

        internal Microsoft.Office.Tools.CustomTaskPane VariablePane { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            VariablePane = this.CustomTaskPanes.Add(VariableControl, "文档内部变量");
            VariablePane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
