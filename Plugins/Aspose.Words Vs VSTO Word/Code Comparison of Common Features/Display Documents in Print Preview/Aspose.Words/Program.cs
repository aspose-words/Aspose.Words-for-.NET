using System;
using System.Windows.Forms;
using Aspose.Words.Rendering;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();

            // Pass the Aspose.Words print document to the Print Preview dialog.
            previewDlg.Document = new AsposeWordsPrintDocument(doc);

            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "TestName.doc";
            previewDlg.WindowState = FormWindowState.Maximized;

            // Show the appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
        }

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
    }
}
