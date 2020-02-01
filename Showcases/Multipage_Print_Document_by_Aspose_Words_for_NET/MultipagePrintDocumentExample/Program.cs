using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Aspose.Words;

namespace MultipagePrintDocumentExample
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            Document doc = new Document(dataDir + "TestFile.docx");

            PrintDialog printDlg = new PrintDialog();
            // Initialize the print dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            // Pass the printer settings from the print dialog to the print document.
            MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Initialize the print preview dialog.
            PrintPreviewDialog previewDlg = new PrintPreviewDialog();

            // Pass the Aspose.Words print document to the print preview dialog.
            previewDlg.Document = awPrintDoc;

            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = doc.OriginalFileName;
            previewDlg.WindowState = FormWindowState.Maximized;

            // Occur whenever the print preview dialog is first displayed.
            previewDlg.Shown += PreviewDlg_Shown;

            // Show the appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
        }

        private static void PreviewDlg_Shown(object sender, EventArgs e)
        {
            // Bring the print preview dialog on top when it is initially displayed.
            ((PrintPreviewDialog)sender).Activate();
        }
    }
}
