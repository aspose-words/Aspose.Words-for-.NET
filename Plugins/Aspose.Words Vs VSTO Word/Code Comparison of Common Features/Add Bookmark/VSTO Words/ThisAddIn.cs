using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Word.Application wordApp = Application;

            string filePath = @"..\..\..\..\..\Sample Files\";

            wordApp.Documents.Open(filePath + "MyDocument.docx");

            Document extendedDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            Bookmark firstParagraph = extendedDocument.Controls.AddBookmark(
                extendedDocument.Paragraphs[1].Range, "FirstParagraph"); 
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
