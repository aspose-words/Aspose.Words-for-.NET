using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string mypath = "";
            Word.Application wordApp = Application;
            wordApp.Documents.Open(mypath + "Add Headers and Footers.doc");

            //Add Header
            foreach (Word.Section section in this.Application.ActiveDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }

            //Add Footer
            foreach (Word.Section wordSection in this.Application.ActiveDocument.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 20;
                footerRange.Text = "Confidential";
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
