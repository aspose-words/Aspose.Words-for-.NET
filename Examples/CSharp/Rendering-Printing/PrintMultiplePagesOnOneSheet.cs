using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words.Rendering;
using Aspose.Words;
using System.Windows.Forms;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class PrintMultiplePagesOnOneSheet
    {
        public static void Run()
        {
            //ExStart:PrintMultiplePagesOnOneSheet
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");
            //ExStart:PrintDialogSettings
            PrintDialog printDlg = new PrintDialog();
            // Initialize the Print Dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;
            //ExEnd:PrintDialogSettings
            // Check if user accepted the print settings and proceed to preview.
            //ExStart:CheckPrintSettings
            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;
            //ExEnd:CheckPrintSettings

            // Pass the printer settings from the dialog to the print document.
            MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
            //ExStart:ActivePrintPreviewDialog
            // Create and configure the the ActivePrintPreviewDialog class.
            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();
            previewDlg.Document = awPrintDoc;
            // Specify additional parameters of the Print Preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.Document.DocumentName = "TestFile.doc";
            previewDlg.WindowState = FormWindowState.Maximized;
            // Show appropriately configured Print Preview dialog.
            previewDlg.ShowDialog();
            //ExEnd:ActivePrintPreviewDialog
            //ExEnd:PrintMultiplePagesOnOneSheet
        }
        
    }
    //ExStart:MultipagePrintDocument
    class MultipagePrintDocument : PrintDocument
    //ExEnd:MultipagePrintDocument
    {
        // The data and state fields of the custom PrintDocument class.
        //ExStart:DataAndStaticFields        
        private readonly Document mDocument;
        private readonly int mPagesPerSheet;
        private readonly bool mPrintPageBorders;
        private Size mPaperSize;
        private int mCurrentPage;
        private int mPageTo;
        //ExEnd:DataAndStaticFields
        /// <summary>
        /// The constructor of the custom PrintDocument class.
        /// </summary> 
        //ExStart:MultipagePrintDocumentConstructor 
        public MultipagePrintDocument(Document document, int pagesPerSheet, bool printPageBorders)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            mDocument = document;
            mPagesPerSheet = pagesPerSheet;
            mPrintPageBorders = printPageBorders;
        }
        //ExEnd:MultipagePrintDocumentConstructor
        /// <summary>
        /// The overridden method OnBeginPrint, which is called before the first page of the document prints.
        /// </summary>
        //ExStart:OnBeginPrint
        protected override void OnBeginPrint(PrintEventArgs e)
        {
            base.OnBeginPrint(e);

            switch (PrinterSettings.PrintRange)
            {
                case PrintRange.AllPages:
                    mCurrentPage = 0;
                    mPageTo = mDocument.PageCount - 1;
                    break;
                case PrintRange.SomePages:
                    mCurrentPage = PrinterSettings.FromPage - 1;
                    mPageTo = PrinterSettings.ToPage - 1;
                    break;
                default:
                    throw new InvalidOperationException("Unsupported print range.");
            }

            // Store the size of the paper, selected by user, taking into account the paper orientation.
            if (PrinterSettings.DefaultPageSettings.Landscape)
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Height,
                                      PrinterSettings.DefaultPageSettings.PaperSize.Width);
            else
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Width,
                                      PrinterSettings.DefaultPageSettings.PaperSize.Height);

        }
        //ExEnd:OnBeginPrint
        /// <summary>
        /// Generates the printed page from the specified number of the document pages.
        /// </summary>
        //ExStart:OnPrintPage
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);

            // Transfer to the point units.
            e.Graphics.PageUnit = GraphicsUnit.Point;
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            // Get the number of the thumbnail placeholders across and down the paper sheet.
            Size thumbCount = GetThumbCount(mPagesPerSheet);

            // Calculate the size of each thumbnail placeholder in points.
            // Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
            SizeF thumbSize = new SizeF(
                HundredthsInchToPoint(mPaperSize.Width) / thumbCount.Width,
                HundredthsInchToPoint(mPaperSize.Height) / thumbCount.Height);

            // Select the number of the last page to be printed on this sheet of paper.
            int pageTo = System.Math.Min(mCurrentPage + mPagesPerSheet - 1, mPageTo);

            // Loop through the selected pages from the stored current page to calculated last page.
            for (int pageIndex = mCurrentPage; pageIndex <= pageTo; pageIndex++)
            {
                // Calculate the column and row indices.
                int columnIdx;
                int rowIdx = System.Math.DivRem(pageIndex - mCurrentPage, thumbCount.Width, out columnIdx);

                // Define the thumbnail location in world coordinates (points in this case).
                float thumbLeft = columnIdx * thumbSize.Width;
                float thumbTop = rowIdx * thumbSize.Height;
                // Render the document page to the Graphics object using calculated coordinates and thumbnail placeholder size.
                // The useful return value is the scale at which the page was rendered.
                float scale = mDocument.RenderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height);

                // Draw the page borders (the page thumbnail could be smaller than the thumbnail placeholder size).
                if (mPrintPageBorders)
                {
                    // Get the real 100% size of the page in points.
                    SizeF pageSize = mDocument.GetPageInfo(pageIndex).SizeInPoints;
                    // Draw the border around the scaled page using the known scale factor.
                    e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, pageSize.Width * scale, pageSize.Height * scale);

                    // Draws the border around the thumbnail placeholder.
                    e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height);
                }
            }

            // Re-calculate next current page and continue with printing if such page resides within the print range.
            mCurrentPage = mCurrentPage + mPagesPerSheet;
            e.HasMorePages = (mCurrentPage <= mPageTo);
        }
        //ExEnd:OnPrintPage
        /// <summary>
        /// Converts hundredths of inches to points.
        /// </summary>
        //ExStart:HundredthsInchToPoint
        private static float HundredthsInchToPoint(float value)
        {
            return (float)ConvertUtil.InchToPoint(value / 100);
        }
        //ExEnd:HundredthsInchToPoint
        /// <summary>
        /// Defines the number of columns and rows depending on the pagesPerSheet number and the page orientation.
        /// </summary>
        //ExStart:GetThumbCount
        private Size GetThumbCount(int pagesPerSheet)
        {
            Size size;
            // Define the number of the columns and rows on the sheet for the Landscape-oriented paper.
            switch (pagesPerSheet)
            {
                case 16: size = new Size(4, 4); break;
                case 9: size = new Size(3, 3); break;
                case 8: size = new Size(4, 2); break;
                case 6: size = new Size(3, 2); break;
                case 4: size = new Size(2, 2); break;
                case 2: size = new Size(2, 1); break;
                default: size = new Size(1, 1); break;
            }
            // Switch the width and height if the paper is in the Portrait orientation.
            if (mPaperSize.Width < mPaperSize.Height)
                return new Size(size.Height, size.Width);
            return size;
        }
        //ExEnd:GetThumbCount
    }
}
