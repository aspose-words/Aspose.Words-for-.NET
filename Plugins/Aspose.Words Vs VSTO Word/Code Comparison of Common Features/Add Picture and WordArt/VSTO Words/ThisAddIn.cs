using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string filePath = @"..\..\..\..\..\Sample Files\";

            Word.Application wordApp = Application;
            wordApp.Documents.Open(filePath + "MyDocument.docx");
            //Add picture to Doc
            this.Application.Selection.InlineShapes.AddPicture(filePath + "Logo.jpg");

            // Add WordArt.
            // Get the left and top position of the current cursor location.
            float leftPosition = (float)this.Application.Selection.Information[
            Word.WdInformation.wdHorizontalPositionRelativeToPage];

            float topPosition = (float)this.Application.Selection.Information[
            Word.WdInformation.wdVerticalPositionRelativeToPage];

            // Call the AddTextEffect method of the Shapes object of the active document (or a different document that you specify).
            this.Application.ActiveDocument.Shapes.AddTextEffect(Office.MsoPresetTextEffect.msoTextEffect29, "test","Arial Black", 24, Office.MsoTriState.msoFalse,
            Office.MsoTriState.msoFalse, leftPosition, topPosition);
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
