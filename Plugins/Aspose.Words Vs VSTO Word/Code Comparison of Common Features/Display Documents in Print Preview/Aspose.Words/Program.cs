using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileName = "YourFileName.docx";
            Document doc = new Document(FileName);

            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();

            // Pass the Aspose.Words print document to the Print Preview dialog.
            previewDlg.Document = doc;
            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "TestName.doc";
            previewDlg.WindowState = FormWindowState.Maximized;
            // Show the appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
        }
    }
}
