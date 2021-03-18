using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string filePath = @"..\..\..\..\..\Sample Files\";
            Word.Application wordApp = Application;

            wordApp.Documents.Open(filePath + "MyDocument.docx");

            int recordCount = 2;
            int i = 0;
            for (i = 0; i <= recordCount; i++)
                wordApp.Selection.WholeStory();
            wordApp.Selection.EndOf();
            wordApp.Selection.InsertFile(filePath + "MyDocument.docx");

            if (i < recordCount)
            {
                wordApp.Selection.Range.InsertBreak(2);
            }
            if (i > 1)
            {
                //wordApp.ActiveDocument.Sections(i).Headers(1).LinkToPrevious = false;
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
