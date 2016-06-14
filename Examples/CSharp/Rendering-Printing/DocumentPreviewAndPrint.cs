
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

using Aspose.Words;
using Aspose.Words.Rendering;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    //ExStart:ActivePrintPreviewDialogClass 
    class ActivePrintPreviewDialog : PrintPreviewDialog
    {
        /// <summary>
        /// Brings the Print Preview dialog on top when it is initially displayed.
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            Activate();
            base.OnShown(e);
        }

    }
    //ExEnd:ActivePrintPreviewDialogClass

    /// <summary>
    /// This project is set to target the x86 platform because the .NET print dialog does not 
    /// seem to show when calling from a 64-bit application.
    /// </summary>
    public class DocumentPreviewAndPrint
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); ;

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            //ExStart:PrintDialog
            //ExId:DocumentPreviewAndPrint_PrintDialog_Creation
            //ExSummary:Creates the print dialog.
            PrintDialog printDlg = new PrintDialog();
            // Initialize the print dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;
            //ExEnd:PrintDialog

            //ExStart:ShowDialog
            //ExId:DocumentPreviewAndPrint_PrintDialog_Check_Result
            //ExSummary:Check if the user accepted the print settings and proceed to preview the document.
            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;
            //ExEnd:ShowDialog

            //ExStart:AsposeWordsPrintDocument
            //ExId:DocumentPreviewAndPrint_AsposeWordsPrintDocument_Creation
            //ExSummary:Creates a special Aspose.Words implementation of the .NET PrintDocument class.
            // Pass the printer settings from the dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
            //ExEnd:AsposeWordsPrintDocument

            //ExStart:ActivePrintPreviewDialog
            //ExId:DocumentPreviewAndPrint_ActivePrintPreviewDialog_Creation
            //ExSummary:Creates an overridden version of the .NET Print Preview dialog to preview the document.
            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();

            // Pass the Aspose.Words print document to the Print Preview dialog.
            previewDlg.Document = awPrintDoc;
            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "TestName.doc";
            previewDlg.WindowState = FormWindowState.Maximized;
            // Show the appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
            //ExEnd:ActivePrintPreviewDialog
        }
    }
}
