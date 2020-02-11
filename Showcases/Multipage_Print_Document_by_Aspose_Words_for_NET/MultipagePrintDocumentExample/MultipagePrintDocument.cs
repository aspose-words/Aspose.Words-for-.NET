using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace MultipagePrintDocumentExample
{
    internal class MultipagePrintDocument : PrintDocument
    {
        /// <summary>
        /// Initializes a new instance of this class.
        /// </summary>
        /// <param name="document">The document to print.</param>
        /// <param name="pagesPerSheet">The number of pages per one sheet.</param>
        /// <param name="printPageBorder">The flag that indicates if the printed page borders are needed.</param>
        public MultipagePrintDocument(Document document, int pagesPerSheet, bool printPageBorder)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            mDocument = document;
            mPagesPerSheet = pagesPerSheet;
            mPrintPageBorder = printPageBorder;
        }

        /// <summary>
        /// Called before the printing starts. Initializes the range of pages to be printed
        /// according to the user selection.
        /// </summary>
        /// <param name="e">The event arguments.</param>
        protected override void OnBeginPrint(PrintEventArgs e)
        {
            base.OnBeginPrint(e);

            switch (PrinterSettings.PrintRange)
            {
                case PrintRange.AllPages:
                    mCurrentPage = 1;
                    mPageTo = mDocument.PageCount;
                    break;
                case PrintRange.SomePages:
                    mCurrentPage = PrinterSettings.FromPage;
                    mPageTo = PrinterSettings.ToPage;
                    break;
                default:
                    throw new InvalidOperationException("Unsupported print range.");
            }

            // Store the page size, selected by user, taking into account the paper orientation.
            if (PrinterSettings.DefaultPageSettings.Landscape)
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Height,
                                      PrinterSettings.DefaultPageSettings.PaperSize.Width);
            else
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Width,
                                      PrinterSettings.DefaultPageSettings.PaperSize.Height);

        }

        private static Size GetThumbCount(int pagesPerSheet)
        {
            switch (pagesPerSheet)
            {
                case 16: return new Size(4, 4);
                case 9: return new Size(3, 3);
                case 8: return new Size(4, 2);
                case 6: return new Size(3, 2);
                case 4: return new Size(2, 2);
                case 2: return new Size(2, 1);
                default: return new Size(1, 1);
            }
        }

        /// <summary>
        /// Called when each page is printed. This method actually renders the page to the graphics object.
        /// </summary>
        /// <param name="e">The event arguments.</param>
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);

            e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            // This gives us how many thumbnails we print across and down.
            Size thumbCount = GetThumbCount(mPagesPerSheet);

            // These are in "display" units (1/100 inch). 
            SizeF thumbSize = new SizeF((float)mPaperSize.Width / thumbCount.Width, (float)mPaperSize.Height / thumbCount.Height);

            mPageTo = (mPageTo + mPagesPerSheet < mPageTo) ? mPageTo + mPagesPerSheet : mPageTo;

            for (int pageIndex = mCurrentPage - 1; pageIndex < mPageTo; pageIndex++)
            {
                int dividend = pageIndex - (mCurrentPage - 1);
                int rowIdx = dividend / thumbCount.Width;
                int columnIdx = dividend % thumbCount.Width;

                // This renders the page in the appropriate location and size given in world coordinates (1/100 inch in our case).
                float thumbLeft = columnIdx * thumbSize.Width;
                float thumbTop = rowIdx * thumbSize.Height;
                // The useful return value is the scale at which the page was rendered.
                float scale = mDocument.RenderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height);

                // This draws the page border (the page could be smaller than the thumbnail size).
                if (mPrintPageBorder)
                {
                    PageInfo pageInfo = mDocument.GetPageInfo(pageIndex);
                    // We know how much the page was scaled so we can draw the border around the scaled page now.
                    e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, WidthInHundredthsInch(pageInfo) * scale, HeightInHundredthsInch(pageInfo) * scale);
                }

                // This draws a border around the thumbnail.
                e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height);
            }

            mCurrentPage = mCurrentPage + mPagesPerSheet;
            e.HasMorePages = (mCurrentPage <= mPageTo);
        }

        /// <summary>
        /// Returns the width of the page in hundredths of an inch.
        /// </summary>
        private int WidthInHundredthsInch(PageInfo pageInfo)
        {
            double widthInInches = pageInfo.WidthInPoints / PointsPerInch;

            return (int)Math.Round(widthInInches * 100);
        }

        /// <summary>
        /// Returns the height of the page in hundredths of an inch.
        /// </summary>
        private int HeightInHundredthsInch(PageInfo pageInfo)
        {
            double heightInInches = pageInfo.HeightInPoints / PointsPerInch;

            return (int)Math.Round(heightInInches * 100);
        }

        public const double PointsPerInch = 72.0;
        private readonly Document mDocument;
        private readonly int mPagesPerSheet;
        private readonly bool mPrintPageBorder;
        private Size mPaperSize = Size.Empty; // Initialized for Java to work.
        private int mCurrentPage;
        private int mPageTo;
    }
}