﻿// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Rendering
{
    [TestFixture]
    public class ExRendering : QaTestsBase
    {
        [Test]
        public void SaveToPdfDefault()
        {
            //ExStart
            //ExFor:Document.Save(String)
            //ExSummary:Converts a whole document to PDF using default options.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            doc.Save(MyDir + "Rendering.SaveToPdfDefault Out.pdf");
            //ExEnd
        }

        [Test]
        public void SaveToPdfWithOutline()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:PdfSaveOptions
            //ExFor:PdfSaveOptions.HeadingsOutlineLevels
            //ExFor:PdfSaveOptions.ExpandedOutlineLevels
            //ExSummary:Converts a whole document to PDF with three levels in the document outline.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();
            options.HeadingsOutlineLevels = 3;
            options.ExpandedOutlineLevels = 1;
            
            doc.Save(MyDir + "Rendering.SaveToPdfWithOutline Out.pdf", options);
            //ExEnd
        }

        [Test]
        public void SaveToPdfStreamOnePage()
        {
            //ExStart
            //ExFor:PdfSaveOptions.PageIndex
            //ExFor:PdfSaveOptions.PageCount
            //ExFor:Document.Save(Stream, SaveOptions)
            //ExSummary:Converts just one page (third page in this example) of the document to PDF.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            using (Stream stream = File.Create(MyDir + "Rendering.SaveToPdfStreamOnePage Out.pdf"))
            {
                PdfSaveOptions options = new PdfSaveOptions();
                options.PageIndex = 2;
                options.PageCount = 1;
                doc.Save(stream, options);
            }
            //ExEnd
        }

        [Test]
        public void SaveToPdfNoCompression()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfSaveOptions.TextCompression
            //ExFor:PdfTextCompression
            //ExSummary:Saves a document to PDF without compression.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();
            options.TextCompression = PdfTextCompression.None;

            doc.Save(MyDir + "Rendering.SaveToPdfNoCompression Out.pdf", options);
            //ExEnd
        }

        [Test]
        public void SaveAsPdf()
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreserveFormFields
            //ExFor:Document.Save(String)
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExFor:Document.Save(String, SaveOptions)
            //ExId:SaveToPdf_NewAPI
            //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
            // Open the document
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Option 1: Save document to file in the PDF format with default options
            doc.Save(MyDir + "Rendering.PdfDefaultOptions Out.pdf");

            // Option 2: Save the document to stream in the PDF format with default options
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Pdf);
            // Rewind the stream position back to the beginning, ready for use
            stream.Seek(0, SeekOrigin.Begin);

            // Option 3: Save document to the PDF format with specified options
            // Render the first page only and preserve form fields as usable controls and not as plain text
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.PageIndex = 0;
            pdfOptions.PageCount = 1;
            pdfOptions.PreserveFormFields = true;
            doc.Save(MyDir + "Rendering.PdfCustomOptions Out.pdf", pdfOptions);
            //ExEnd
        }

        [Test]
        public void SaveAsXps()
        {
            //ExStart
            //ExFor:XpsSaveOptions
            //ExFor:XpsSaveOptions.#ctor
            //ExFor:Document.Save(String)
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExFor:Document.Save(String, SaveOptions)
            //ExId:SaveToXps_NewAPI
            //ExSummary:Shows how to save a document to the Xps format using the Save method and the XpsSaveOptions class.
            // Open the document
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");
            // Save document to file in the Xps format with default options
            doc.Save(MyDir + "Rendering.XpsDefaultOptions Out.xps");

            // Save document to stream in the Xps format with default options
            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Xps);
            // Rewind the stream position back to the beginning, ready for use
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to file in the Xps format with specified options
            // Render the first page only
            XpsSaveOptions xpsOptions = new XpsSaveOptions();
            xpsOptions.PageIndex = 0;
            xpsOptions.PageCount = 1;
            doc.Save(MyDir + "Rendering.XpsCustomOptions Out.xps", xpsOptions);
            //ExEnd
        }

        [Test]
        public void SaveAsImage()
        {
            //ExStart
            //ExFor:ImageSaveOptions.#ctor
            //ExFor:Document.Save(String)
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExFor:Document.Save(String, SaveOptions)
            //ExId:SaveToImage_NewAPI
            //ExSummary:Shows how to save a document to the Jpeg format using the Save method and the ImageSaveOptions class.
            // Open the document
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");
            // Save as a Jpeg image file with default options
            doc.Save(MyDir + "Rendering.JpegDefaultOptions Out.jpg");

            // Save document to stream as a Jpeg with default options
            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Jpeg);
            // Rewind the stream position back to the beginning, ready for use
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to a Jpeg image with specified options.
            // Render the third page only and set the jpeg quality to 80%
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
            // to signal what type of image to save as.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            imageOptions.PageIndex = 2;
            imageOptions.PageCount = 1;
            imageOptions.JpegQuality = 80;
            doc.Save(MyDir + "Rendering.JpegCustomOptions Out.jpg", imageOptions);
            //ExEnd
        }

        [Test]
        public void SaveToTiffDefault()
        {
            //ExStart
            //ExFor:Document.Save(String)
            //ExSummary:Converts a whole document into a multipage TIFF file using default options.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            doc.Save(MyDir + "Rendering.SaveToTiffDefault Out.tiff");
            //ExEnd
        }

        [Test]
        public void SaveToTiffCompression()
        {
            //ExStart
            //ExFor:TiffCompression
            //ExFor:ImageSaveOptions.TiffCompression
            //ExFor:ImageSaveOptions.PageIndex
            //ExFor:ImageSaveOptions.PageCount
            //ExFor:Document.Save(String, SaveOptions)
            //ExSummary:Converts a page of a Word document into a TIFF image and uses the CCITT compression.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.TiffCompression = TiffCompression.Ccitt3;
            options.PageIndex = 0;
            options.PageCount = 1;

            doc.Save(MyDir + "Rendering.SaveToTiffCompression Out.tiff", options);
            //ExEnd
        }

        [Test]
        public void SaveToImageResolution()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.Resolution
            //ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.Resolution = 300;
            options.PageCount = 1;

            doc.Save(MyDir + "Rendering.SaveToImageResolution Out.png", options);
            //ExEnd
        }

        [Test]
        public void SaveToEmf()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Emf);
            options.PageCount = 1;

            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(MyDir + "Rendering.SaveToEmf." + i.ToString() + " Out.emf", options);
            }
            //ExEnd
        }

        [Test]
        public void SaveToImageJpegQuality()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.JpegQuality
            //ExSummary:Converts a page of a Word document into JPEG images of different qualities.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            // Try worst quality.
            options.JpegQuality = 0;
            doc.Save(MyDir + "Rendering.SaveToImageJpegQuality0 Out.jpeg", options);

            // Try best quality.
            options.JpegQuality = 100;
            doc.Save(MyDir + "Rendering.SaveToImageJpegQuality100 Out.jpeg", options);
            //ExEnd
        }

        [Test]
        public void SaveToImagePaperColor()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.PaperColor
            //ExSummary:Renders a page of a Word document into an image with transparent or coloured background.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);

            imgOptions.PaperColor = Color.Transparent;
            doc.Save(MyDir + "Rendering.SaveToImagePaperColorTransparent Out.png", imgOptions);

            imgOptions.PaperColor = Color.LightCoral;
            doc.Save(MyDir + "Rendering.SaveToImagePaperColorCoral Out.png", imgOptions);
            //ExEnd
        }

        [Test]
        public void SaveToImageStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExSummary:Saves a document page as a BMP image into a stream.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Bmp);

            // Rewind the stream and create a .NET image from it.
            stream.Position = 0;

            // Read the stream back into an image.
            Image image = Image.FromStream(stream); 
            //ExEnd
        }

        [Test]
        public void UpdatePageLayout()
        {
            //ExStart
            //ExFor:StyleCollection.Item(String)
            //ExFor:SectionCollection.Item(Int32)
            //ExFor:Document.UpdatePageLayout
            //ExSummary:Shows when to request page layout of the document to be recalculated.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Saving a document to PDF or to image or printing for the first time will automatically
            // layout document pages and this information will be cached inside the document.
            doc.Save(MyDir + "Rendering.UpdatePageLayout1 Out.pdf");

            // Modify the document in any way.
            doc.Styles["Normal"].Font.Size = 6;
            doc.Sections[0].PageSetup.Orientation = Aspose.Words.Orientation.Landscape;

            // In the current version of Aspose.Words, modifying the document does not automatically rebuild 
            // the cached page layout. If you want to save to PDF or render a modified document again,
            // you need to manually request page layout to be updated.
            doc.UpdatePageLayout();

            doc.Save(MyDir + "Rendering.UpdatePageLayout2 Out.pdf");
            //ExEnd
        }

        [Test]
        public void UpdateFieldsBeforeRendering()
        {
            //ExStart
            //ExFor:Document.UpdateFields
            //ExId:UpdateFieldsBeforeRendering
            //ExSummary:Shows how to update all fields before rendering a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // This updates all fields in the document.
            doc.UpdateFields();

            doc.Save(MyDir + "Rendering.UpdateFields Out.pdf");
            //ExEnd
        }

        [Test, Explicit]
        public void Print()
        {
            //ExStart
            //ExFor:Document.Print
            //ExSummary:Prints the whole document to the default printer.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            doc.Print();
            //ExEnd
        }

        [Test, Explicit]
        public void PrintToNamedPrinter()
        {
            //ExStart
            //ExFor:Document.Print(String)
            //ExSummary:Prints the whole document to a specified printer.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            doc.Print("KONICA MINOLTA magicolor 2400W");
            //ExEnd
        }

        [Test, Explicit]
        public void PrintRange()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings)
            //ExSummary:Prints a range of pages.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PrinterSettings printerSettings = new PrinterSettings();
            // Page numbers in the .NET printing framework are 1-based.
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            doc.Print(printerSettings);
            //ExEnd
        }

        [Test, Explicit]
        public void PrintRangeWithDocumentName()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings, String)
            //ExSummary:Prints a range of pages along with the name of the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PrinterSettings printerSettings = new PrinterSettings();
            // Page numbers in the .NET printing framework are 1-based.
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            doc.Print(printerSettings, "My Print Document.doc");
            //ExEnd
        }

        [Test, Explicit]
        public void PreviewAndPrint()
        {
            //ExStart
            //ExFor:AsposeWordsPrintDocument
            //ExSummary:Shows the Print dialog that allows selecting the printer and page range to print with. Then brings up the print preview from which you can preview the document and choose to print or close.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PrintPreviewDialog previewDlg = new PrintPreviewDialog();
            // Show non-modal first is a hack for the print preview form to show on top.
            previewDlg.Show();

            // Initialize the Print Dialog with the number of pages in the document.
            PrintDialog printDlg = new PrintDialog();
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;

            // Create the Aspose.Words' implementation of the .NET print document 
            // and pass the printer settings from the dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Hide and invalidate preview is a hack for print preview to show on top.
            previewDlg.Hide();
            previewDlg.PrintPreviewControl.InvalidatePreview();

            // Pass the Aspose.Words' print document to the .NET Print Preview dialog.
            previewDlg.Document = awPrintDoc;

            previewDlg.ShowDialog();
            //ExEnd
        }

        [Test]
        public void RenderToScale()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExFor:Document.GetPageInfo
            //ExFor:PageInfo
            //ExFor:PageInfo.GetSizeInPixels
            //ExSummary:Renders a page of a Word document into a bitmap using a specified zoom factor.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PageInfo pageInfo = doc.GetPageInfo(0);

            // Let's say we want the image at 50% zoom.
            const float MyScale = 0.50f;

            // Let's say we want the image at this resolution.
            const float MyResolution = 200.0f;

            Size pageSize = pageInfo.GetSizeInPixels(MyScale, MyResolution);

            using (Bitmap img = new Bitmap(pageSize.Width, pageSize.Height))
            {
                img.SetResolution(MyResolution, MyResolution);

                using (Graphics gr = Graphics.FromImage(img))
                {
                    // You can apply various settings to the Graphics object.
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Fill the page background.
                    gr.FillRectangle(Brushes.White, 0, 0, pageSize.Width, pageSize.Height);

                    // Render the page using the zoom.
                    doc.RenderToScale(0, gr, 0, 0, MyScale);
                }

                img.Save(MyDir + "Rendering.RenderToScale Out.png");
            }
            //ExEnd
        }

        [Test]
        public void RenderToSize()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Render to a bitmap at a specified location and size.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            using (Bitmap bmp = new Bitmap(700, 700))
            {
                // User has some sort of a Graphics object. In this case created from a bitmap.
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    // The user can specify any options on the Graphics object including
                    // transform, antialiasing, page units, etc.
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Let's say we want to fit the page into a 3" x 3" square on the screen so use inches as units.
                    gr.PageUnit = GraphicsUnit.Inch;

                    // The output should be offset 0.5" from the edge and rotated.
                    gr.TranslateTransform(0.5f, 0.5f);
                    gr.RotateTransform(10);

                    // This is our test rectangle.
                    gr.DrawRectangle(new Pen(Color.Black, 3f / 72f), 0f, 0f, 3f, 3f);

                    // User specifies (in world coordinates) where on the Graphics to render and what size.
                    float returnedScale = doc.RenderToSize(0, gr, 0f, 0f, 3f, 3f);

                    // This is the calculated scale factor to fit 297mm into 3".
                    Console.WriteLine("The image was rendered at {0:P0} zoom.", returnedScale);


                    // One more example, this time in millimiters.
                    gr.PageUnit = GraphicsUnit.Millimeter;

                    gr.ResetTransform();

                    // Move the origin 10mm 
                    gr.TranslateTransform(10, 10);

                    // Apply both scale transform and page scale for fun.
                    gr.ScaleTransform(0.5f, 0.5f);
                    gr.PageScale = 2f;

                    // This is our test rectangle.
                    gr.DrawRectangle(new Pen(Color.Black, 1), 90, 10, 50, 100);

                    // User specifies (in world coordinates) where on the Graphics to render and what size.
                    doc.RenderToSize(1, gr, 90, 10, 50, 100);


                    bmp.Save(MyDir + "Rendering.RenderToSize Out.png");
                }
            }
            //ExEnd
        }

        [Test]
        public void createThumbnails()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages.

            // The user opens or builds a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // This defines the number of columns to display the thumbnails in.
            const int thumbColumns = 2;

            // Calculate the required number of rows for thumbnails.
            // We can now get the number of pages in the document.
            int remainder;
            int thumbRows = Math.DivRem(doc.PageCount, thumbColumns, out remainder);
            if (remainder > 0)
                thumbRows++;

            // Lets say I want thumbnails to be of this zoom.
            const float scale = 0.25f;

            // For simplicity lets pretend all pages in the document are of the same size, 
            // so we can use the size of the first page to calculate the size of the thumbnail.
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails.
            int imgWidth = thumbSize.Width * thumbColumns;
            int imgHeight = thumbSize.Height * thumbRows;

            using (Bitmap img = new Bitmap(imgWidth, imgHeight))
            {
                // The user has to provides a Graphics object to draw on.
                // The Graphics object can be created from a bitmap, from a metafile, printer or window.
                using (Graphics gr = Graphics.FromImage(img))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Fill the "paper" with white, otherwise it will be transparent.
                    gr.FillRectangle(new SolidBrush(Color.White), 0, 0, imgWidth, imgHeight);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int columnIdx;
                        int rowIdx = Math.DivRem(pageIndex, thumbColumns, out columnIdx);

                        // Specify where we want the thumbnail to appear.
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        SizeF size = doc.RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);

                        // Draw the page rectangle.
                        gr.DrawRectangle(Pens.Black, thumbLeft, thumbTop, size.Width, size.Height);
                    }

                    img.Save(MyDir + "Rendering.Thumbnails Out.png");
                }
            }
            //ExEnd
        }

        //ExStart
        //ExFor:PageInfo.GetDotNetPaperSize
        //ExFor:PageInfo.Landscape
        //ExSummary:Shows how to implement your own .NET PrintDocument to completely customize printing of Aspose.Words documents.
        [Test, Explicit] //ExSkip
        public void CustomPrint()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Create an instance of our own PrintDocument.
            MyPrintDocument printDoc = new MyPrintDocument(doc);
            // Specify the page range to print.
            printDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printDoc.PrinterSettings.FromPage = 1;
            printDoc.PrinterSettings.ToPage = 1;
            
            // Print our document.
            printDoc.Print();
        }

        /// <summary>
        /// The way to print in the .NET Framework is to implement a class derived from PrintDocument.
        /// This class is an example on how to implement custom printing of an Aspose.Words document.
        /// It selects an appropriate paper size, orientation and paper tray when printing.
        /// </summary>
        public class MyPrintDocument : PrintDocument
        {
            public MyPrintDocument(Aspose.Words.Document document)
            {
                mDocument = document;
            }

            /// <summary>
            /// Called before the printing starts. 
            /// </summary>
            protected override void OnBeginPrint(PrintEventArgs e)
            {
                base.OnBeginPrint(e);

                // Initialize the range of pages to be printed according to the user selection.
                switch (PrinterSettings.PrintRange)
                {
                    case System.Drawing.Printing.PrintRange.AllPages:
                        mCurrentPage = 1;
                        mPageTo = mDocument.PageCount;
                        break;
                    case System.Drawing.Printing.PrintRange.SomePages:
                        mCurrentPage = PrinterSettings.FromPage;
                        mPageTo = PrinterSettings.ToPage;
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported print range.");
                }
            }

            /// <summary>
            /// Called before each page is printed. 
            /// </summary>
            protected override void OnQueryPageSettings(QueryPageSettingsEventArgs e)
            {
                base.OnQueryPageSettings(e);

                // A single Word document can have multiple sections that specify pages with different sizes, 
                // orientation and paper trays. This code is called by the .NET printing framework before 
                // each page is printed and we get a chance to specify how the page is to be printed.
                PageInfo pageInfo = mDocument.GetPageInfo(mCurrentPage - 1);
                e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(PrinterSettings.PaperSizes);
                // MS Word stores the paper source (printer tray) for each section as a printer-specfic value.
                // To obtain the correct tray value you will need to use the RawKindValue returned
                // by .NET for your printer.
                e.PageSettings.PaperSource.RawKind = pageInfo.PaperTray;
                e.PageSettings.Landscape = pageInfo.Landscape;
            }

            /// <summary>
            /// Called for each page to render it for printing. 
            /// </summary>
            protected override void OnPrintPage(PrintPageEventArgs e)
            {
                base.OnPrintPage(e);

                // Aspose.Words rendering engine creates a page that is drawn from the 0,0 of the paper,
                // but there is some hard margin in the printer and the .NET printing framework
                // renders from there. We need to offset by that hard margin.

                // In .NET 1.1 the hard margin is not available programmatically, lets hardcode to about 4mm.
                float hardOffsetX = 20;
                float hardOffsetY = 20;

                // This is in .NET 2.0 only. Uncomment when needed.
//                float hardOffsetX = e.PageSettings.HardMarginX;
//                float hardOffsetY = e.PageSettings.HardMarginY;

                int pageIndex = mCurrentPage - 1;
                mDocument.RenderToScale(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f);

                mCurrentPage++;
                e.HasMorePages = (mCurrentPage <= mPageTo);
            }

            private readonly Aspose.Words.Document mDocument;
            private int mCurrentPage;
            private int mPageTo;
        }
        //ExEnd

        [Test, Explicit]
        public void WritePageInfo()
        {
            //ExStart
            //ExFor:PageInfo
            //ExFor:PageInfo.PaperSize
            //ExFor:PageInfo.PaperTray
            //ExFor:PageInfo.Landscape
            //ExFor:PageInfo.WidthInPoints
            //ExFor:PageInfo.HeightInPoints
            //ExSummary:Retrieves page size and orientation information for every page in a Word document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");
            
            Console.WriteLine("Document \"{0}\" contains {1} pages.", doc.OriginalFileName, doc.PageCount);

            for (int i = 0; i < doc.PageCount; i++)
            {
                PageInfo pageInfo = doc.GetPageInfo(i);
                Console.WriteLine(
                    "Page {0}. PaperSize:{1} ({2:F0}x{3:F0}pt), Orientation:{4}, PaperTray:{5}", 
                    i + 1,
                    pageInfo.PaperSize,
                    pageInfo.WidthInPoints,
                    pageInfo.HeightInPoints,
                    pageInfo.Landscape ? "Landscape" : "Portrait",
                    pageInfo.PaperTray);
            }
            //ExEnd
        }

        [Test]
        public void SetTrueTypeFontsFolder()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] fontSources = FontSettings.GetFontsSources();

            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolder(String, Boolean)
            //ExId:SetFontsFolderCustomFolder
            //ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering or embedding of fonts.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for 
            // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
            // FontSettings.SetFontSources instead.
            FontSettings.SetFontsFolder(@"C:\MyFonts\", false);

            doc.Save(MyDir + "Rendering.SetFontsFolder Out.pdf");
            //ExEnd

            // Restore the original sources used to search for fonts.
            FontSettings.SetFontsSources(fontSources);
        }

        [Test]
        public void SetFontsFoldersMultipleFolders()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] fontSources = FontSettings.GetFontsSources();

            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
            //ExId:SetFontsFoldersMultipleFolders
            //ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering or embedding fonts.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for 
            // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
            // FontSettings.SetFontSources instead.
            FontSettings.SetFontsFolders(new string[] {@"C:\MyFonts\", @"D:\Misc\Fonts\"}, true);

            doc.Save(MyDir + "Rendering.SetFontsFolders Out.pdf");
            //ExEnd

            // Restore the original sources used to search for fonts.
            FontSettings.SetFontsSources(fontSources);
        }

        [Test]
        public void SetFontsFoldersSystemAndCustomFolder()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] origFontSources = FontSettings.GetFontsSources();

            //ExStart
            //ExFor:FontSettings            
            //ExFor:FontSettings.GetFontsSources()
            //ExFor:FontSettings.SetFontsSources()
            //ExId:SetFontsFoldersSystemAndCustomFolder
            //ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders as well as a custom defined folder when scanning for fonts.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
            // We add this array to a new ArrayList to make adding or removing font entries much easier.
            ArrayList fontSources = new ArrayList(FontSettings.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts. 
            FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

            // Add the custom folder which contains our fonts to the list of existing font sources.
            fontSources.Add(folderFontSource);

            // Convert the Arraylist of source back into a primitive array of FontSource objects.
            FontSourceBase[] updatedFontSources = (FontSourceBase[])fontSources.ToArray(typeof(FontSourceBase));

            // Apply the new set of font sources to use.
            FontSettings.SetFontsSources(updatedFontSources);

            doc.Save(MyDir + "Rendering.SetFontsFolders Out.pdf");
            //ExEnd

            // Verify that font sources are set correctly.
            Assert.IsInstanceOf(typeof(SystemFontSource), FontSettings.GetFontsSources()[0]); // The first source should be a system font source.
            Assert.IsInstanceOf(typeof(FolderFontSource), FontSettings.GetFontsSources()[1]); // The second source should be our folder font source.

            FolderFontSource folderSource = ((FolderFontSource)FontSettings.GetFontsSources()[1]);
            Assert.AreEqual(@"C:\MyFonts\", folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);

            // Restore the original sources used to search for fonts.
            FontSettings.SetFontsSources(origFontSources);
        }

        [Test]
        public void SaveToPdfWithJpegImageCompression()
        {
            //ExStart
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExId:SaveToPdfJpegImageCompression
            //ExSummary:Demonstrates how to save images to PDF using JPEG encoding to decrease file size.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();

            // Use JPEG compression at 50% quality to reduce file size.
            options.ImageCompression = PdfImageCompression.Jpeg;
            options.JpegQuality = 50;

            doc.Save(MyDir + "Rendering.JpegImageCompression Out.pdf", options);
            //ExEnd
        }

        [Test]
        public void SetDefaultFontName()
        {
            //ExStart
            //ExFor:FontSettings.DefaultFontName
            //ExId:SetDefaultFontName
            //ExSummary:Demonstrates how to specify what font to substitute for a missing font during rendering.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
            FontSettings.DefaultFontName = "Arial Unicode MS";

            // Now the set default font is used in place of any missing fonts during any rendering calls.
            doc.Save(MyDir + "Rendering.SetDefaultFont Out.pdf");
            doc.Save(MyDir + "Rendering.SetDefaultFont Out.xps");
            //ExEnd
        }

        [Test]
        public void RecieveFontSubstitutionNotification()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] origFontSources = FontSettings.GetFontsSources();

            //ExStart
            //ExFor:IWarningCallback
            //ExFor:SaveOptions.WarningCallback
            //ExId:FontSubstitutionNotification
            //ExSummary:Demonstrates how to recieve notifications of font substitutions by using IWarningCallback.
            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            // We can choose the default font to use in the case of any missing fonts.
            FontSettings.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
            FontSettings.SetFontsFolder(string.Empty, false);

            // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            // We assign the callback to the appropriate save options class. In this case, we are going to save to PDF
            // so we create a PdfSaveOptions class and assign the callback there.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.WarningCallback = callback;

            // Pass the save options along with the save path to the save method.
            doc.Save(MyDir + "Rendering.MissingFontNotification Out.pdf", saveOptions);
            //ExEnd
            
            Assert.Greater(callback.mFontWarnings.Count, 0);
            Assert.True(callback.mFontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.mFontWarnings[0].Description.Contains("has not been found"));

            // Restore default fonts. 
            FontSettings.SetFontsSources(origFontSources);
        }

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:SaveOptions.WarningCallback
        //ExId:FontSubstitutionWarningCallback
        //ExSummary:Demonstrates how to implement the IWarningCallback to be notified of any font substitution during document save.
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document procssing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                    mFontWarnings.Warning(info); //ExSkip
                }
            }

            public WarningInfoCollection mFontWarnings = new WarningInfoCollection(); //ExSkip
        }
        //ExEnd

        [Test]
        public void RecieveFontSubstitutionUpdatePageLayout()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] origFontSources = FontSettings.GetFontsSources();

            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            // We can choose the default font to use in the case of any missing fonts.
            FontSettings.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
            FontSettings.SetFontsFolder(string.Empty, false);

            //ExStart
            //ExId:FontSubstitutionUpdatePageLayout
            //ExSummary:Demonstrates how IWarningCallback will still recieve warning notifcations even if UpdatePageLayout is called before document save.
            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
            // are stored until the document save and then sent to the appropriate WarningCallback.
            doc.UpdatePageLayout();

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.WarningCallback = callback;

            // Even though the document was rendered previously, any save warnings are notified to the user during document save.
            doc.Save(MyDir + "Rendering.FontsNotificationUpdatePageLayout Out.pdf", saveOptions);
            //ExEnd

            Assert.Greater(callback.mFontWarnings.Count, 0);
            Assert.True(callback.mFontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.mFontWarnings[0].Description.Contains("has not been found"));

            // Restore default fonts. 
            FontSettings.SetFontsSources(origFontSources);
        }

        [Test]
        public void EmbedFullFontsInPdf()
        {
            //ExStart
            //ExFor:PdfSaveOptions.#ctor
            //ExFor:PdfSaveOptions.EmbedFullFonts
            //ExId:EmbedFullFonts
            //ExSummary:Demonstrates how to set Aspose.Words to embed full fonts in the output PDF document.
            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // each time a document is rendered.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document.
            doc.Save(MyDir + "Rendering.EmbedFullFonts Out.pdf");
            //ExEnd
        }

        [Test]
        public void SubsetFontsInPdf()
        {
            //ExStart
            //ExFor:PdfSaveOptions.EmbedFullFonts
            //ExId:Subset
            //ExSummary:Demonstrates how to set Aspose.Words to subset fonts in the output PDF.
            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;

            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // in the document are included in the PDF fonts.
            doc.Save(MyDir + "Rendering.SubsetFonts Out.pdf");
            //ExEnd
        }

        [Test]
        public void DisableEmbeddingStandardWindowsFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.EmbedStandardWindowsFonts
            //ExId:EmbedStandardWindowsFonts
            //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedStandardWindowsFonts = false;

            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(MyDir + "Rendering.DisableEmbedWindowsFonts Out.pdf");
            //ExEnd
        }

        [Test]
        public void DisableEmbeddingCoreFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseCoreFonts
            //ExId:DisableUseOfCoreFonts
            //ExSummary:Shows how to set Aspose.Words to avoid embedding core fonts and let the reader subsuite PDF Type 1 fonts instead.
            // Load the document to render.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");

            // To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
            PdfSaveOptions options = new PdfSaveOptions();
            options.UseCoreFonts = true;

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            doc.Save(MyDir + "Rendering.DisableEmbedWindowsFonts Out.pdf");
            //ExEnd
        }

        [Test]
        public void SetPdfEncryptionPermissions()
        {
            //ExStart
            //ExFor:PdfEncryptionDetails.#ctor
            //ExFor:PdfSaveOptions.EncryptionDetails
            //ExFor:PdfEncryptionDetails.Permissions
            //ExFor:PdfEncryptionAlgorithm
            //ExFor:PdfPermissions
            //ExFor:PdfEncryptionDetails
            //ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.doc");
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            
            // Create encryption details and set owner password.
            PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails(string.Empty, "password", PdfEncryptionAlgorithm.RC4_128);

            // Start by disallowing all permissions.
            encryptionDetails.Permissions = PdfPermissions.DisallowAll;

            // Extend permissions to allow editing or modifying annotations.
            encryptionDetails.Permissions = PdfPermissions.ModifyAnnotations | PdfPermissions.DocumentAssembly;
            saveOptions.EncryptionDetails = encryptionDetails;

            // Render the document to PDF format with the specified permissions.
            doc.Save(MyDir + "Rendering.SpecifyPermissions Out.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void SetPdfNumeralFormat()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Rendering.NumeralFormat.doc");
            //ExStart
            //ExFor:PdfSaveOptions.NumeralFormat
            //ExSummary:Demonstrates how to set the numeral format used when saving to PDF.
            PdfSaveOptions options = new PdfSaveOptions();
            options.NumeralFormat = NumeralFormat.Context;
            //ExEnd

            doc.Save(MyDir + "Rendering.NumeralFormat Out.pdf", options);
        }
    }
}
