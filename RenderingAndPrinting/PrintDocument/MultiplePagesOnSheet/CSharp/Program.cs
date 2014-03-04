//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

using Aspose.Words;

namespace MultiplePagesOnSheetExample
{
    /// <summary>
    /// A subclass of the .NET Print Preview dialog. This extension only is used only to work around the .NET PrintPreviewDialog
    /// class not appearing in front of other windows by default.
    /// </summary>
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

    /// <summary>
    /// This project is set to target the x86 platform because the .NET print dialog does not 
    /// seem to show when calling from a 64-bit application.
    /// </summary>
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //ExStart
            //ExId:MultiplePagesOnSheet_PrintAndPreview
            //ExSummary:The usage of the MultipagePrintDocument for Previewing and Printing.
            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            PrintDialog printDlg = new PrintDialog();
            // Initialize the Print Dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            // Check if user accepted the print settings and proceed to preview.
            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;

            // Pass the printer settings from the dialog to the print document.
            MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

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
            //ExEnd
        }
    }

    //ExStart
    //ExId:MultiplePagesOnSheet_PrintDocument
    //ExSummary:The custom PrintDocument class.
    class MultipagePrintDocument : PrintDocument
    //ExEnd
    {
        /// <summary>
        /// Initializes a new instance of this class.
        /// </summary>
        /// <param name="document">The document to print.</param>
        /// <param name="pagesPerSheet">The number of pages per one sheet.</param>
        /// <param name="printPageBorders">The flag that indicates if the printed page borders are rendered.</param>
        //ExStart
        //ExId:MultiplePagesOnSheet_Constructor
        //ExSummary:The constructor of the custom PrintDocument class.
        public MultipagePrintDocument(Document document, int pagesPerSheet, bool printPageBorders)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            mDocument = document;
            mPagesPerSheet = pagesPerSheet;
            mPrintPageBorders = printPageBorders;
        }
        //ExEnd

        /// <summary>
        /// Called before the printing starts. Initializes the range of pages to be printed
        /// according to the user's selection.
        /// </summary>
        /// <param name="e">The event arguments.</param>
        //ExStart
        //ExId:MultiplePagesOnSheet_OnBeginPrint
        //ExSummary:The overridden method OnBeginPrint, which is called before the first page of the document prints.
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
        //ExEnd

        /// <summary>
        /// Converts the pagesPerSheet number into the number of columns and rows.
        /// </summary>
        /// <param name="pagesPerSheet">The number of the pages to be printed on the one sheet of paper.</param>
        //ExStart
        //ExId:MultiplePagesOnSheet_GetThumbCount
        //ExSummary:Defines the number of columns and rows depending on the pagesPerSheet number and the page orientation.
        private Size GetThumbCount(int pagesPerSheet)
        {
            Size size;
            // Define the number of the columns and rows on the sheet for the Landscape-oriented paper.
            switch (pagesPerSheet)
            {
                case   16: size = new Size(4, 4); break;
                case    9: size = new Size(3, 3); break;
                case    8: size = new Size(4, 2); break;
                case    6: size = new Size(3, 2); break;
                case    4: size = new Size(2, 2); break;
                case    2: size = new Size(2, 1); break;
                default  : size = new Size(1, 1); break;
            }
            // Switch the width and height if the paper is in the Portrait orientation.
            if (mPaperSize.Width < mPaperSize.Height)
                return new Size(size.Height, size.Width);
            return size;
        }
        //ExEnd

        /// <summary>
        /// Converts 1/100 inch into 1/72 inch (points).
        /// </summary>
        /// <param name="value">The 1/100 inch value to convert.</param>
        //ExStart
        //ExId:MultiplePagesOnSheet_HundredthsInchToPoint
        //ExSummary:Converts hundredths of inches to points.
        private static float HundredthsInchToPoint(float value)
        {
            return (float)ConvertUtil.InchToPoint(value / 100);
        }
        //ExEnd
        
        /// <summary>
        /// Called when each page is printed. This method actually renders the page to the graphics object.
        /// </summary>
        /// <param name="e">The event arguments.</param>
        //ExStart
        //ExId:MultiplePagesOnSheet_OnPrintPage
        //ExSummary:Generates the printed page from the specified number of the document pages.
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);

            // Transfer to the point units.
            e.Graphics.PageUnit = GraphicsUnit.Point;
            e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            // Get the number of the thumbnail placeholders across and down the paper sheet.
            Size thumbCount = GetThumbCount(mPagesPerSheet);

            // Calculate the size of each thumbnail placeholder in points.
            // Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
            SizeF thumbSize = new SizeF(
                HundredthsInchToPoint(mPaperSize.Width) / thumbCount.Width,
                HundredthsInchToPoint(mPaperSize.Height) / thumbCount.Height);

            // Select the number of the last page to be printed on this sheet of paper.
            int pageTo = Math.Min(mCurrentPage + mPagesPerSheet - 1, mPageTo);

            // Loop through the selected pages from the stored current page to calculated last page.
            for (int pageIndex = mCurrentPage; pageIndex <= pageTo; pageIndex++)
            {
                // Calculate the column and row indices.
                int columnIdx;
                int rowIdx = Math.DivRem(pageIndex - mCurrentPage, thumbCount.Width, out columnIdx);

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
        //ExEnd

        //ExStart
        //ExId:MultiplePagesOnSheet_Fields
        //ExSummary:The data and state fields of the custom PrintDocument class.
        private readonly Document mDocument;
        private readonly int mPagesPerSheet;
        private readonly bool mPrintPageBorders;
        private Size mPaperSize;
        private int mCurrentPage;
        private int mPageTo;
        //ExEnd
    }
}