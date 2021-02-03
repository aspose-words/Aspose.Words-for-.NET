using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string mypath = "Document.docx";
            Word.Application wordApp = Application;
            wordApp.Documents.Open(mypath);

            foreach (Table table in this.Application.ActiveDocument.Tables)
            {
                table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;
            }
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
