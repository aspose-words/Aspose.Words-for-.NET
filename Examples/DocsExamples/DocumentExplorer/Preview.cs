using System.Drawing.Printing;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace DocumentExplorer
{
	/// <summary>
	/// Provides a utility method to print preview and print an Aspose.Words document.
	/// </summary>
	internal class Preview
	{
	    /// <summary>
	    /// No ctor.
	    /// </summary>
	    private Preview()
	    {
	    }
	    
	    /// <summary>
        /// A utility method to print preview and print an Aspose.Words document.
        /// </summary>
        internal static void Execute(Document document)
        {
            // This operation can take some time (for the first page) so we set the Cursor to WaitCursor.
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

	        PrintPreviewDialog previewDlg = new PrintPreviewDialog();

            // Initialize the Print Dialog with the number of pages in the document.
            PrintDialog printDlg = new PrintDialog();
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings = new PrinterSettings();
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = document.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = document.PageCount;

            // Restore cursor.
            Cursor.Current = cursor;

            // Interesting, but PrintDialog will not show and will always return cancel
            // If you run this application in 64-bit mode.
	        if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;

            // Create the Aspose.Words' implementation of the .NET print document 
            // And pass the printer settings from the dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(document);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Pass the Aspose.Words' print document to the .NET Print Preview dialog.
            previewDlg.Document = awPrintDoc;

            previewDlg.ShowDialog();
        }

    }
}