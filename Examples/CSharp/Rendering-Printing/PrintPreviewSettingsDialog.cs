using Aspose.Words.Rendering;
using System;
using System.Windows.Forms;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class PrintPreviewSettingsDialog
    {
        public static void Run()
        {
            // ExStart:PrintPreviewSettingsDialog
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            Document doc = new Document(dataDir + "TestFile.doc");

            PrintDialog printDlg = new PrintDialog();

            // Initialize the print dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            // Сheck if the user accepted the print settings and whether to proceed to document preview.
            if (printDlg.ShowDialog() != DialogResult.OK)
                return;

            // Create a special Aspose.Words implementation of the .NET PrintDocument class.
            // Pass the printer settings from the print dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Initialize the print preview dialog.
            PrintPreviewDialog previewDlg = new PrintPreviewDialog();

            // Pass the Aspose.Words print document to the print preview dialog.
            previewDlg.Document = awPrintDoc;

            // Specify additional parameters of the print preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = doc.OriginalFileName;
            previewDlg.WindowState = FormWindowState.Maximized;

            // Occur whenever the print preview dialog is first displayed.
            previewDlg.Shown += PreviewDlg_Shown;

            // Show the appropriately configured print preview dialog.
            previewDlg.ShowDialog();
            // ExEnd:PrintPreviewSettingsDialog
        }

        // ExStart:PrintPreviewSettingsDialogEvent
        private static void PreviewDlg_Shown(object sender, EventArgs e)
        {
            // Bring the print preview dialog on top when it is initially displayed.
            ((PrintPreviewDialog)sender).Activate();
        }
        // ExEnd:PrintPreviewSettingsDialogEvent
    }
}
