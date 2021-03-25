using Word = Microsoft.Office.Interop.Word;

namespace VSTO_Words
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            object noReset = false;
            object password = "MyPassword";
            object useIRM = false;
            object enforceStyleLock = false;

            this.Application.ActiveDocument.Protect(Word.WdProtectionType.wdAllowOnlyReading,
                ref noReset, ref password, ref useIRM, ref enforceStyleLock);
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
