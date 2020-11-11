// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Collections;
using System.IO;
using System.Linq;
using Aspose.Pdf.Text;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using FolderFontSource = Aspose.Words.Fonts.FolderFontSource;
using SystemFontSource = Aspose.Words.Fonts.SystemFontSource;
#if NET462 || JAVA
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing.Text;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExRendering : ApiExampleBase
    {
        [Test]
        public void SaveToPdfStreamOnePage()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.PageIndex
            //ExFor:FixedPageSaveOptions.PageCount
            //ExFor:Document.Save(Stream, SaveOptions)
            //ExSummary:Shows how to convert only some of the pages in a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            using (Stream stream = File.Create(ArtifactsDir + "Rendering.SaveToPdfStreamOnePage.pdf"))
            {
                // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
                // to modify the way in which that method converts the document to .PDF.
                PdfSaveOptions options = new PdfSaveOptions();

                // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
                options.PageIndex = 1;

                // Set the "PageCount" to "1" to render only one page of the document,
                // starting from the page that the "PageIndex" property specified.
                options.PageCount = 1;
                
                // This document will contain one page starting from page two, which means it will only contain the second page.
                doc.Save(stream, options);
            }
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "Rendering.SaveToPdfStreamOnePage.pdf");

            Assert.AreEqual(1, pdfDocument.Pages.Count);

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            Assert.AreEqual("Page 2.", textFragmentAbsorber.Text);
#endif
        }

        [TestCase(PdfTextCompression.None)]
        [TestCase(PdfTextCompression.Flate)]
        public void TextCompression(PdfTextCompression pdfTextCompression)
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfSaveOptions.TextCompression
            //ExFor:PdfTextCompression
            //ExSummary:Shows how to apply text compression when saving a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < 100; i++)
                builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
            // compression to text when we save the document to PDF.
            // Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
            // to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
            options.TextCompression = pdfTextCompression;

            doc.Save(ArtifactsDir + "Rendering.TextCompression.pdf", options);

            switch (pdfTextCompression)
            {
                case PdfTextCompression.None:
                    Assert.That(60000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.TextCompression.pdf").Length));
                    TestUtil.FileContainsString("5 0 obj\r\n<</Length 9 0 R>>stream", ArtifactsDir + "Rendering.TextCompression.pdf"); //ExSkip
                    break;
                case PdfTextCompression.Flate:
                    Assert.That(30000, Is.AtLeast(new FileInfo(ArtifactsDir + "Rendering.TextCompression.pdf").Length));
                    TestUtil.FileContainsString("5 0 obj\r\n<</Length 9 0 R/Filter /FlateDecode>>stream", ArtifactsDir + "Rendering.TextCompression.pdf"); //ExSkip
                    break;
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PreserveFormFields(bool preserveFormFields)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreserveFormFields
            //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
            // Open the document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please select a fruit: ");

            // Insert a combo box which will allow a user to choose an option from a collection of strings.
            builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
            // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
            // their current values, and display them as plain text in the output PDF.
            pdfOptions.PreserveFormFields = preserveFormFields;

            doc.Save(ArtifactsDir + "Rendering.PreserveFormFields.pdf", pdfOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "Rendering.PreserveFormFields.pdf");

            Assert.AreEqual(1, pdfDocument.Pages.Count);

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            if (preserveFormFields)
            {
                Assert.AreEqual("Please select a fruit: ", textFragmentAbsorber.Text);
                TestUtil.FileContainsString("10 0 obj\r\n" +
                                            "<</Type /Annot/Subtype /Widget/P 4 0 R/FT /Ch/F 4/Rect [168.39199829 707.35101318 217.87442017 722.64007568]/Ff 131072/T(þÿ\0M\0y\0C\0o\0m\0b\0o\0B\0o\0x)/Opt " +
                                            "[(þÿ\0A\0p\0p\0l\0e) (þÿ\0B\0a\0n\0a\0n\0a) (þÿ\0C\0h\0e\0r\0r\0y) ]/V(þÿ\0A\0p\0p\0l\0e)/DA(0 g /FAAABC 12 Tf )/AP<</N 11 0 R>>>>", 
                    ArtifactsDir + "Rendering.PreserveFormFields.pdf");
            }
            else
            {
                Assert.AreEqual("Please select a fruit: Apple", textFragmentAbsorber.Text);
                Assert.Throws<AssertionException>(() =>
                {
                    TestUtil.FileContainsString("/Widget", 
                        ArtifactsDir + "Rendering.PreserveFormFields.pdf");
                });
            }
#endif
        }

        [Test]
        public void SaveAsXps()
        {
            //ExStart
            //ExFor:XpsSaveOptions
            //ExFor:XpsSaveOptions.#ctor
            //ExFor:XpsSaveOptions.OutlineOptions
            //ExFor:XpsSaveOptions.SaveFormat
            //ExSummary:Shows how to limit the level of headings that will appear in the outline of a saved XPS document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 1.2.1");
            builder.Writeln("Heading 1.2.2");

            // Create an "XpsSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .XPS.
            XpsSaveOptions saveOptions = new XpsSaveOptions();
            
            Assert.AreEqual(SaveFormat.Xps, saveOptions.SaveFormat);

            // The output XPS document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
            // The last two headings we have inserted above will not appear.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 2;

            doc.Save(ArtifactsDir + "Rendering.SaveAsXps.xps", saveOptions);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SaveAsXpsBookFold(bool renderTextAsBookfold)
        {
            //ExStart
            //ExFor:XpsSaveOptions.#ctor(SaveFormat)
            //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create an "XpsSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .XPS.
            XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.Xps);

            // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            // in the output XPS in a way that helps us use it to make a booklet.
            // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
            xpsOptions.UseBookFoldPrintingSettings = true;

            // If we are rendering the document as a booklet, we must set the "MultiplePages"
            // properties of all page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            if (renderTextAsBookfold)
                foreach (Section s in doc.Sections)
                {
                    s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
                }

            // Once we print this document, we can turn it into a booklet by stacking the pages
            // in the order they come out of the printer and then folding down the middle
            doc.Save(ArtifactsDir + $"Rendering.SaveAsXpsBookFold.{renderTextAsBookfold}.xps", xpsOptions);
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
            //ExSummary:Shows how to save a document to the JPEG format using the Save method and the ImageSaveOptions class.
            // Open the document
            Document doc = new Document(MyDir + "Rendering.docx");
            // Save as a JPEG image file with default options
            doc.Save(ArtifactsDir + "Rendering.SaveAsImage.DefaultJpgOptions.jpg");

            // Save document to stream as a JPEG with default options
            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Jpeg);
            // Rewind the stream position back to the beginning, ready for use
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to a JPEG image with specified options
            // Render the third page only and set the JPEG quality to 80%
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
            // to signal what type of image to save as
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            imageOptions.PageIndex = 2;
            imageOptions.PageCount = 1;
            imageOptions.JpegQuality = 80;
            doc.Save(ArtifactsDir + "Rendering.SaveAsImage.CustomJpgOptions.jpg", imageOptions);
            //ExEnd
        }

        [Test, Category("SkipMono")]
        public void SaveToTiffDefault()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Rendering.SaveToTiffDefault.tiff");
        }

        [Test, Category("SkipMono")]
        public void SaveToTiffCompression()
        {
            //ExStart
            //ExFor:TiffCompression
            //ExFor:ImageSaveOptions.TiffCompression
            //ExFor:ImageSaveOptions.PageIndex
            //ExFor:ImageSaveOptions.PageCount
            //ExSummary:Converts a page of a Word document into a TIFF image and uses the CCITT compression.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                PageIndex = 0,
                PageCount = 1
            };

            doc.Save(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff", options);
            //ExEnd
        }

        [Test]
        public void SaveToImageResolution()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.Resolution
            //ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,
                PageCount = 1
            };

            doc.Save(ArtifactsDir + "Rendering.SaveToImageResolution.png", options);
            //ExEnd
        }

        [Test, Category("SkipMono")]
        public void SaveToEmf()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions
            //ExFor:Document.Save(String, SaveOptions)
            //ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Emf) { PageCount = 1 };

            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(ArtifactsDir + "Rendering.SaveToEmf." + i + ".emf", options);
            }
            //ExEnd
        }

        [Test]
        public void SaveToImageJpegQuality()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.JpegQuality
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.JpegQuality
            //ExSummary:Converts a page of a Word document into JPEG images of different qualities.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            // Try worst quality
            saveOptions.JpegQuality = 0;
            doc.Save(ArtifactsDir + "Rendering.SaveToImageJpegQuality.0.jpeg", saveOptions);

            // Try best quality
            saveOptions.JpegQuality = 100;
            doc.Save(ArtifactsDir + "Rendering.SaveToImageJpegQuality.100.jpeg", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveToImagePaperColor()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.PaperColor
            //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);

            imgOptions.PaperColor = Color.Transparent;
            doc.Save(ArtifactsDir + "Rendering.SaveToImagePaperColor.Transparent.png", imgOptions);

            imgOptions.PaperColor = Color.LightCoral;
            doc.Save(ArtifactsDir + "Rendering.SaveToImagePaperColor.Coral.png", imgOptions);
            //ExEnd
        }

        #if NET462 || JAVA
        [Test]
        public void SaveToImageStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExSummary:Saves a document page as a BMP image into a stream.
            Document doc = new Document(MyDir + "Rendering.docx");

            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Bmp);

            // Rewind the stream and create a .NET image from it
            stream.Position = 0;

            // Read the stream back into an image
            using (Image image = Image.FromStream(stream))
            {
                // ...Do something
            }
            //ExEnd
        }

        [Test]
        public void RenderToSize()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Render to a bitmap at a specified location and size.
            Document doc = new Document(MyDir + "Rendering.docx");
            
            using (Bitmap bmp = new Bitmap(700, 700))
            {
                // User has some sort of a Graphics object. In this case created from a bitmap
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    // The user can specify any options on the Graphics object including
                    // transform, anti-aliasing, page units, etc.
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // If we want to fit the page into a 3" x 3" square on the screen, we will need to set the measurement units to inches
                    gr.PageUnit = GraphicsUnit.Inch;

                    // The output should be offset 0.5" from the edge and rotated
                    gr.TranslateTransform(0.5f, 0.5f);
                    gr.RotateTransform(10);

                    // This is our test rectangle
                    gr.DrawRectangle(new Pen(Color.Black, 3f / 72f), 0f, 0f, 3f, 3f);

                    // User specifies (in world coordinates) where on the Graphics to render and what size
                    float returnedScale = doc.RenderToSize(0, gr, 0f, 0f, 3f, 3f);

                    // This is the calculated scale factor to fit 297mm into 3"
                    Console.WriteLine("The image was rendered at {0:P0} zoom.", returnedScale);

                    // One more example, this time in millimeters
                    gr.PageUnit = GraphicsUnit.Millimeter;

                    gr.ResetTransform();

                    // Move the origin 10mm 
                    gr.TranslateTransform(10, 10);

                    // Apply both scale transform and page scale for fun
                    gr.ScaleTransform(0.5f, 0.5f);
                    gr.PageScale = 2f;

                    // This is our test rectangle
                    gr.DrawRectangle(new Pen(Color.Black, 1), 90, 10, 50, 100);

                    // User specifies (in world coordinates) where on the Graphics to render and what size
                    doc.RenderToSize(1, gr, 90, 10, 50, 100);

                    bmp.Save(ArtifactsDir + "Rendering.RenderToSize.png");
                }
            }
            //ExEnd
        }

        [Test]
        public void Thumbnails()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages.
            // The user opens or builds a document
            Document doc = new Document(MyDir + "Rendering.docx");

            // This defines the number of columns to display the thumbnails in
            const int thumbColumns = 2;

            // Calculate the required number of rows for thumbnails
            // We can now get the number of pages in the document
            int thumbRows = Math.DivRem(doc.PageCount, thumbColumns, out int remainder);
            if (remainder > 0)
                thumbRows++;

            // Define a zoom factor for the thumbnails 
            const float scale = 0.25f;

            // We can use the size of the first page to calculate the size of the thumbnail,
            // assuming that all pages in the document are of the same size
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails
            int imgWidth = thumbSize.Width * thumbColumns;
            int imgHeight = thumbSize.Height * thumbRows;
            
            using (Bitmap img = new Bitmap(imgWidth, imgHeight))
            {
                // The Graphics object, which we will draw on, can be created from a bitmap, metafile, printer, or window
                using (Graphics gr = Graphics.FromImage(img))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Fill the "paper" with white, otherwise it will be transparent
                    gr.FillRectangle(new SolidBrush(Color.White), 0, 0, imgWidth, imgHeight);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int rowIdx = Math.DivRem(pageIndex, thumbColumns, out int columnIdx);

                        // Specify where we want the thumbnail to appear
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        SizeF size = doc.RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);

                        // Draw the page rectangle
                        gr.DrawRectangle(Pens.Black, thumbLeft, thumbTop, size.Width, size.Height);
                    }

                    img.Save(ArtifactsDir + "Rendering.Thumbnails.png");
                }
            }
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void CustomPrint()
        {
            //ExStart
            //ExFor:PageInfo.GetDotNetPaperSize
            //ExFor:PageInfo.Landscape
            //ExSummary:Shows how to implement your own .NET PrintDocument to completely customize printing of Aspose.Words documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create an instance of our own PrintDocument
            MyPrintDocument printDoc = new MyPrintDocument(doc);
            // Specify the page range to print
            printDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printDoc.PrinterSettings.FromPage = 1;
            printDoc.PrinterSettings.ToPage = 1;

            // Print our document.
            printDoc.Print();
        }

        /// <summary>
        /// The way to print in the .NET Framework is to implement a class derived from PrintDocument.
        /// This class is an example on how to implement custom printing of an Aspose.Words document.
        /// It selects an appropriate paper size, orientation, and paper tray when printing.
        /// </summary>
        public class MyPrintDocument : PrintDocument
        {
            public MyPrintDocument(Document document)
            {
                mDocument = document;
            }

            /// <summary>
            /// Called before the printing starts. 
            /// </summary>
            protected override void OnBeginPrint(PrintEventArgs e)
            {
                base.OnBeginPrint(e);

                // Initialize the range of pages to be printed according to the user selection
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
                // orientations, and paper trays. This code is called by the .NET printing framework before 
                // each page is printed and we get a chance to specify how the page is to be printed
                PageInfo pageInfo = mDocument.GetPageInfo(mCurrentPage - 1);
                e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(PrinterSettings.PaperSizes);
                // MS Word stores the paper source (printer tray) for each section as a printer-specific value
                // To obtain the correct tray value you will need to use the RawKindValue returned
                // by .NET for your printer
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
                // renders from there. We need to offset by that hard margin

                // In .NET 1.1 the hard margin is not available programmatically, set it to approximately 4mm
                float hardOffsetX = 20;
                float hardOffsetY = 20;

                // This is in .NET 2.0 only. Uncomment when needed
                // float hardOffsetX = e.PageSettings.HardMarginX;
                // float hardOffsetY = e.PageSettings.HardMarginY;

                int pageIndex = mCurrentPage - 1;
                mDocument.RenderToScale(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f);

                mCurrentPage++;
                e.HasMorePages = (mCurrentPage <= mPageTo);
            }

            private readonly Document mDocument;
            private int mCurrentPage;
            private int mPageTo;
        }
        //ExEnd

        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void PrintPageInfo()
        {
            //ExStart
            //ExFor:PageInfo
            //ExFor:PageInfo.GetSizeInPixels(Single, Single, Single)
            //ExFor:PageInfo.GetSpecifiedPrinterPaperSource(PaperSourceCollection, PaperSource)
            //ExFor:PageInfo.HeightInPoints
            //ExFor:PageInfo.Landscape
            //ExFor:PageInfo.PaperSize
            //ExFor:PageInfo.PaperTray
            //ExFor:PageInfo.SizeInPoints
            //ExFor:PageInfo.WidthInPoints
            //ExSummary:Shows how to print page size and orientation information for every page in a Word document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // The first section has 2 pages
            // We will assign a different printer paper tray to each one, whose number will match a kind of paper source
            // These sources and their Kinds will vary depending on the installed printer driver
            PrinterSettings.PaperSourceCollection paperSources = new PrinterSettings().PaperSources;

            doc.FirstSection.PageSetup.FirstPageTray = paperSources[0].RawKind;
            doc.FirstSection.PageSetup.OtherPagesTray = paperSources[1].RawKind;

            Console.WriteLine("Document \"{0}\" contains {1} pages.", doc.OriginalFileName, doc.PageCount);

            float scale = 1.0f;
            float dpi = 96;

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Each page has a PageInfo object, whose index is the respective page's number
                PageInfo pageInfo = doc.GetPageInfo(i);

                // Print the page's orientation and dimensions
                Console.WriteLine($"Page {i + 1}:");
                Console.WriteLine($"\tOrientation:\t{(pageInfo.Landscape ? "Landscape" : "Portrait")}");
                Console.WriteLine($"\tPaper size:\t\t{pageInfo.PaperSize} ({pageInfo.WidthInPoints:F0}x{pageInfo.HeightInPoints:F0}pt)");
                Console.WriteLine($"\tSize in points:\t{pageInfo.SizeInPoints}");
                Console.WriteLine($"\tSize in pixels:\t{pageInfo.GetSizeInPixels(1.0f, 96)} at {scale * 100}% scale, {dpi} dpi");

                // Paper source tray information
                Console.WriteLine($"\tTray:\t{pageInfo.PaperTray}");
                PaperSource source = pageInfo.GetSpecifiedPrinterPaperSource(paperSources, paperSources[0]);
                Console.WriteLine($"\tSuitable print source:\t{source.SourceName}, kind: {source.Kind}");
            }
            //ExEnd
        }

        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void PrinterSettingsContainer()
        {
            //ExStart
            //ExFor:PrinterSettingsContainer
            //ExFor:PrinterSettingsContainer.#ctor(PrinterSettings)
            //ExFor:PrinterSettingsContainer.DefaultPageSettingsPaperSource
            //ExFor:PrinterSettingsContainer.PaperSizes
            //ExFor:PrinterSettingsContainer.PaperSources
            //ExSummary:Shows how to access and list your printer's paper sources and sizes.
            // The PrinterSettingsContainer contains a PrinterSettings object,
            // which contains unique data for different printer drivers
            PrinterSettingsContainer container = new PrinterSettingsContainer(new PrinterSettings());

            // You can find the printer's list of paper sources here
            Console.WriteLine($"{container.PaperSources.Count} printer paper sources:");
            foreach (PaperSource paperSource in container.PaperSources)
            {
                bool isDefault = container.DefaultPageSettingsPaperSource.SourceName == paperSource.SourceName;
                Console.WriteLine($"\t{paperSource.SourceName}, " +
                                  $"RawKind: {paperSource.RawKind} {(isDefault ? "(Default)" : "")}");
            }

            // You can find the list of PaperSizes that can be sent to the printer here
            // Both the PrinterSource and PrinterSize contain a "RawKind" attribute,
            // which equates to a paper type listed on the PaperSourceKind enum
            // If the list of PaperSources contains a PaperSource with the same RawKind as that of the page being printed,
            // the page will be printed by the paper source and on the appropriate paper size by the printer
            // Otherwise, the printer will default to the source designated by DefaultPageSettingsPaperSource 
            Console.WriteLine($"{container.PaperSizes.Count} paper sizes:");
            foreach (System.Drawing.Printing.PaperSize paperSize in container.PaperSizes)
            {
                Console.WriteLine($"\t{paperSize}, RawKind: {paperSize.RawKind}");
            }
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void Print()
        {
            //ExStart
            //ExFor:Document.Print
            //ExSummary:Prints the whole document to the default printer.
            Document doc = new Document(MyDir + "Document.docx");
            doc.Print();
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void PrintToNamedPrinter()
        {
            //ExStart
            //ExFor:Document.Print(String)
            //ExSummary:Prints the whole document to a specified printer.
            Document doc = new Document(MyDir + "Document.docx");
            doc.Print("KONICA MINOLTA magicolor 2400W");
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void PrintRange()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings)
            //ExSummary:Prints a range of pages.
            Document doc = new Document(MyDir + "Rendering.docx");

            PrinterSettings printerSettings = new PrinterSettings();
            // Page numbers in the .NET printing framework are 1-based
            printerSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            doc.Print(printerSettings);
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void PrintRangeWithDocumentName()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings, String)
            //ExSummary:Prints a range of pages along with the name of the document.
            Document doc = new Document(MyDir + "Rendering.docx");

            PrinterSettings printerSettings = new PrinterSettings();
            // Page numbers in the .NET printing framework are 1-based
            printerSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            doc.Print(printerSettings, "Rendering.PrintRangeWithDocumentName.docx");
            //ExEnd
        }

        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void PreviewAndPrint()
        {
            //ExStart
            //ExFor:AsposeWordsPrintDocument.#ctor(Document)
            //ExFor:AsposeWordsPrintDocument.CachePrinterSettings
            //ExSummary:Shows the Print dialog that allows selecting the printer and page range to print with. Then brings up the print preview from which you can preview the document and choose to print or close.
            Document doc = new Document(MyDir + "Rendering.docx");

            PrintPreviewDialog previewDlg = new PrintPreviewDialog();
            // Show non-modal first is a hack for the print preview form to show on top
            previewDlg.Show();

            // Initialize the Print Dialog with the number of pages in the document
            PrintDialog printDlg = new PrintDialog();
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;

            // Create the Aspose.Words' implementation of the .NET print document 
            // and pass the printer settings from the dialog to the print document
            // Use 'CachePrinterSettings' to reduce time of first call of Print() method
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
            awPrintDoc.CachePrinterSettings();

            // Hide and invalidate preview is a hack for print preview to show on top
            previewDlg.Hide();
            previewDlg.PrintPreviewControl.InvalidatePreview();

            // Pass the Aspose.Words' print document to the .NET Print Preview dialog
            previewDlg.Document = awPrintDoc;

            previewDlg.ShowDialog();
            //ExEnd
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        public void RenderToSizeNetStandard2()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Render to a bitmap at a specified location and size (.NetStandard 2.0).
            Document doc = new Document(MyDir + "Rendering.docx");
            
            using (SKBitmap bitmap = new SKBitmap(700, 700))
            {
                // User has some sort of a Graphics object. In this case created from a bitmap
                using (SKCanvas canvas = new SKCanvas(bitmap))
                {
                    // Apply scale transform
                    canvas.Scale(70);

                    // The output should be offset 0.5" from the edge and rotated
                    canvas.Translate(0.5f, 0.5f);
                    canvas.RotateDegrees(10);

                    // This is our test rectangle
                    SKRect rect = new SKRect(0f, 0f, 3f, 3f);
                    canvas.DrawRect(rect, new SKPaint
                    {
                        Color = SKColors.Black,
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = 3f / 72f
                    });

                    // User specifies (in world coordinates) where on the Graphics to render and what size
                    float returnedScale = doc.RenderToSize(0, canvas, 0f, 0f, 3f, 3f);

                    Console.WriteLine("The image was rendered at {0:P0} zoom.", returnedScale);

                    // One more example, this time in millimeters
                    canvas.ResetMatrix();

                    // Apply scale transform
                    canvas.Scale(5);

                    // Move the origin 10mm 
                    canvas.Translate(10, 10);

                    // This is our test rectangle
                    rect = new SKRect(0, 0, 50, 100);
                    rect.Offset(90, 10);
                    canvas.DrawRect(rect, new SKPaint
                    {
                        Color = SKColors.Black,
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = 1
                    });

                    // User specifies (in world coordinates) where on the Graphics to render and what size
                    doc.RenderToSize(0, canvas, 90, 10, 50, 100);

                    using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "Rendering.RenderToSizeNetStandard2.png"))
                    {
                        bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                    }
                }
            }            
            //ExEnd
        }

        [Test]
        public void CreateThumbnailsNetStandard2()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages (.NetStandard 2.0).
            // The user opens or builds a document
            Document doc = new Document(MyDir + "Rendering.docx");

            // This defines the number of columns to display the thumbnails in
            const int thumbColumns = 2;

            // Calculate the required number of rows for thumbnails
            // We can now get the number of pages in the document
            int thumbRows = Math.DivRem(doc.PageCount, thumbColumns, out int remainder);
            if (remainder > 0)
                thumbRows++;

            // Define a zoom factor for the thumbnails 
            const float scale = 0.25f;

            // We can use the size of the first page to calculate the size of the thumbnail,
            // assuming that all pages in the document are of the same size
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails
            int imgWidth = thumbSize.Width * thumbColumns;
            int imgHeight = thumbSize.Height * thumbRows;

            using (SKBitmap bitmap = new SKBitmap(imgWidth, imgHeight))
            {
                // The Graphics object, which we will draw on, can be created from a bitmap, metafile, printer, or window
                using (SKCanvas canvas = new SKCanvas(bitmap))
                {
                    // Fill the "paper" with white, otherwise it will be transparent
                    canvas.Clear(SKColors.White);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int rowIdx = Math.DivRem(pageIndex, thumbColumns, out int columnIdx);

                        // Specify where we want the thumbnail to appear
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        SizeF size = doc.RenderToScale(pageIndex, canvas, thumbLeft, thumbTop, scale);

                        // Draw the page rectangle
                        SKRect rect = new SKRect(0, 0, size.Width, size.Height);
                        rect.Offset(thumbLeft, thumbTop);
                        canvas.DrawRect(rect, new SKPaint
                        {
                            Color = SKColors.Black,
                            Style = SKPaintStyle.Stroke
                        });
                    }

                    using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "Rendering.CreateThumbnailsNetStandard2.png"))
                    {
                        bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                    }
                }
            }            
            //ExEnd
        }
#endif

        [Test]
        public void UpdatePageLayout()
        {
            //ExStart
            //ExFor:StyleCollection.Item(String)
            //ExFor:SectionCollection.Item(Int32)
            //ExFor:Document.UpdatePageLayout
            //ExSummary:Shows when to request page layout of the document to be recalculated.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Saving a document to PDF or to image or printing for the first time will automatically
            // layout document pages and this information will be cached inside the document
            doc.Save(ArtifactsDir + "Rendering.UpdatePageLayout.1.pdf");

            // Modify the document in any way
            doc.Styles["Normal"].Font.Size = 6;
            doc.Sections[0].PageSetup.Orientation = Aspose.Words.Orientation.Landscape;

            // In the current version of Aspose.Words, modifying the document does not automatically rebuild 
            // the cached page layout. If you want to save to PDF or render a modified document again,
            // you need to manually request page layout to be updated
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "Rendering.UpdatePageLayout.2.pdf");
            //ExEnd
        }

        [Test]
        public void SetTrueTypeFontsFolder()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] fontSources = FontSettings.DefaultInstance.GetFontsSources();

            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolder(String, Boolean)
            //ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering or embedding of fonts.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Note that this setting will override any default font sources that are being searched by default
            // Now only these folders will be searched for fonts when rendering or embedding fonts
            // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
            // FontSettings.SetFontSources instead
            FontSettings.DefaultInstance.SetFontsFolder(@"C:\MyFonts\", false);

            doc.Save(ArtifactsDir + "Rendering.SetTrueTypeFontsFolder.pdf");
            //ExEnd

            // Restore the original sources used to search for fonts
            FontSettings.DefaultInstance.SetFontsSources(fontSources);
        }

        [Test]
        public void SetFontsFoldersMultipleFolders()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] fontSources = FontSettings.DefaultInstance.GetFontsSources();

            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
            //ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering or embedding fonts.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Note that this setting will override any default font sources that are being searched by default
            // Now only these folders will be searched for fonts when rendering or embedding fonts
            // To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and 
            // FontSettings.SetFontSources instead
            FontSettings.DefaultInstance.SetFontsFolders(new string[] { @"C:\MyFonts\", @"D:\Misc\Fonts\" }, true);

            doc.Save(ArtifactsDir + "Rendering.SetFontsFoldersMultipleFolders.pdf");
            //ExEnd

            // Restore the original sources used to search for fonts
            FontSettings.DefaultInstance.SetFontsSources(fontSources);
        }

        [Test]
        public void SetFontsFoldersSystemAndCustomFolder()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] origFontSources = FontSettings.DefaultInstance.GetFontsSources();

            //ExStart
            //ExFor:FontSettings            
            //ExFor:FontSettings.GetFontsSources()
            //ExFor:FontSettings.SetFontsSources()
            //ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders as well as a custom defined folder when scanning for fonts.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Retrieve the array of environment-dependent font sources that are searched by default
            // For example, this will contain a "Windows\Fonts\" source on a Windows machines
            // We add this array to a new ArrayList to make adding or removing font entries much easier
            ArrayList fontSources = new ArrayList(FontSettings.DefaultInstance.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
            FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

            // Add the custom folder which contains our fonts to the list of existing font sources
            fontSources.Add(folderFontSource);

            // Convert the ArrayList of source back into a primitive array of FontSource objects
            FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.ToArray(typeof(FontSourceBase));

            // Apply the new set of font sources to use
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            doc.Save(ArtifactsDir + "Rendering.SetFontsFoldersSystemAndCustomFolder.pdf");
            //ExEnd

            // The first source should be a system font source
            Assert.That(FontSettings.DefaultInstance.GetFontsSources()[0], Is.InstanceOf(typeof(SystemFontSource))); 
            // The second source should be our folder font source
            Assert.That(FontSettings.DefaultInstance.GetFontsSources()[1], Is.InstanceOf(typeof(FolderFontSource))); 
            
            FolderFontSource folderSource = ((FolderFontSource) FontSettings.DefaultInstance.GetFontsSources()[1]);
            Assert.AreEqual(@"C:\MyFonts\", folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);

            // Restore the original sources used to search for fonts
            FontSettings.DefaultInstance.SetFontsSources(origFontSources);
        }

        [Test]
        public void SetSpecifyFontFolder()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(FontsDir, false);

            // Using load options
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;

            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);

            FolderFontSource folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[0]);

            Assert.AreEqual(FontsDir, folderSource.FolderPath);
            Assert.False(folderSource.ScanSubfolders);
        }

        [Test]
        public void SetFontSubstitutes()
        {
            //ExStart
            //ExFor:Document.FontSettings
            //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
            //ExSummary:Shows how to define alternative fonts if original does not exist
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Times New Roman", new string[] { "Slab", "Arvo" });
            //ExEnd
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.FontSettings = fontSettings;

            // Check that font source are default
            FontSourceBase[] fontSource = doc.FontSettings.GetFontsSources();
            Assert.AreEqual("SystemFonts", fontSource[0].Type.ToString());

            Assert.AreEqual("Times New Roman", doc.FontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName);

            string[] alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Times New Roman").ToArray();
            Assert.AreEqual(new string[] { "Slab", "Arvo" }, alternativeFonts);
        }

        [Test]
        public void SetSpecifyFontFolders()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolders(new string[] { FontsDir, @"C:\Windows\Fonts\" }, true);

            // Using load options
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);

            FolderFontSource folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[0]);
            Assert.AreEqual(FontsDir, folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);

            folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[1]);
            Assert.AreEqual(@"C:\Windows\Fonts\", folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);
        }

        [Test]
        public void AddFontSubstitutes()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Slab", new string[] { "Times New Roman", "Arial" });
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arvo", new string[] { "Open Sans", "Arial" });

            Document doc = new Document(MyDir + "Rendering.docx");
            doc.FontSettings = fontSettings;

            string[] alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Slab").ToArray();
            Assert.AreEqual(new string[] { "Times New Roman", "Arial" }, alternativeFonts);

            alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Arvo").ToArray();
            Assert.AreEqual(new string[] { "Open Sans", "Arial" }, alternativeFonts);
        }

        [Test]
        public void SetDefaultFontName()
        {
            //ExStart
            //ExFor:DefaultFontSubstitutionRule.DefaultFontName
            //ExSummary:Demonstrates how to specify what font to substitute for a missing font during rendering.
            Document doc = new Document(MyDir + "Rendering.docx");

            // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";

            // Now the set default font is used in place of any missing fonts during any rendering calls
            doc.Save(ArtifactsDir + "Rendering.SetDefaultFontName.pdf");
            doc.Save(ArtifactsDir + "Rendering.SetDefaultFontName.xps");
            //ExEnd
        }

        [Test]
        public void UpdatePageLayoutWarnings()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] origFontSources = FontSettings.DefaultInstance.GetFontsSources();

            // Load the document to render
            Document doc = new Document(MyDir + "Document.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            // We can choose the default font to use in the case of any missing fonts
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
            FontSettings.DefaultInstance.SetFontsFolder(string.Empty, false);

            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
            // are stored until the document save and then sent to the appropriate WarningCallback
            doc.UpdatePageLayout();

            // Even though the document was rendered previously, any save warnings are notified to the user during document save
            doc.Save(ArtifactsDir + "Rendering.UpdatePageLayoutWarnings.pdf");
            
            Assert.That(callback.FontWarnings.Count, Is.GreaterThan(0));
            Assert.True(callback.FontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.FontWarnings[0].Description.Contains("has not been found"));

            // Restore default fonts
            FontSettings.DefaultInstance.SetFontsSources(origFontSources);
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                    FontWarnings.Warning(info); //ExSkip
                }
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection(); //ExSkip
        }

        [Test]
        public void EmbedFullFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.#ctor
            //ExFor:PdfSaveOptions.EmbedFullFonts
            //ExSummary:Demonstrates how to set Aspose.Words to embed full fonts in the output PDF document.
            // Load the document to render
            Document doc = new Document(MyDir + "Rendering.docx");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true
            // The property below can be changed each time a document is rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document
            doc.Save(ArtifactsDir + "Rendering.EmbedFullFonts.pdf", options);
            //ExEnd
        }

        [Test]
        public void SubsetFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.EmbedFullFonts
            //ExSummary:Demonstrates how to set Aspose.Words to subset fonts in the output PDF.
            // Load the document to render
            Document doc = new Document(MyDir + "Rendering.docx");

            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;

            // The output PDF will contain subsets of the fonts in the document
            // Only the glyphs used in the document are included in the PDF fonts
            doc.Save(ArtifactsDir + "Rendering.SubsetFonts.pdf", options);
            //ExEnd
        }

        [Test]
        public void DisableEmbedWindowsFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.FontEmbeddingMode
            //ExFor:PdfFontEmbeddingMode
            //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
            // Load the document to render
            Document doc = new Document(MyDir + "Rendering.docx");

            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone;

            // The output PDF will be saved without embedding standard windows fonts
            doc.Save(ArtifactsDir + "Rendering.DisableEmbedWindowsFonts.pdf", options);
            //ExEnd
        }

        [Test]
        public void DisableEmbedCoreFonts()
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseCoreFonts
            //ExSummary:Shows how to set Aspose.Words to avoid embedding core fonts and let the reader substitute PDF Type 1 fonts instead.
            // Load the document to render
            Document doc = new Document(MyDir + "Rendering.docx");

            // To disable embedding of core fonts and substitute PDF type 1 fonts set UseCoreFonts to true
            PdfSaveOptions options = new PdfSaveOptions();
            options.UseCoreFonts = true;

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            doc.Save(ArtifactsDir + "Rendering.DisableEmbedCoreFonts.pdf", options);
            //ExEnd
        }

        [Test]
        public void EncryptionPermissions()
        {
            //ExStart
            //ExFor:PdfEncryptionDetails.#ctor
            //ExFor:PdfSaveOptions.EncryptionDetails
            //ExFor:PdfEncryptionDetails.Permissions
            //ExFor:PdfEncryptionDetails.EncryptionAlgorithm
            //ExFor:PdfEncryptionDetails.OwnerPassword
            //ExFor:PdfEncryptionDetails.UserPassword
            //ExFor:PdfEncryptionAlgorithm
            //ExFor:PdfPermissions
            //ExFor:PdfEncryptionDetails
            //ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Create encryption details and set owner password
            PdfEncryptionDetails encryptionDetails =
                new PdfEncryptionDetails("password", string.Empty, PdfEncryptionAlgorithm.RC4_128);

            // Start by disallowing all permissions
            encryptionDetails.Permissions = PdfPermissions.DisallowAll;

            // Extend permissions to allow editing or modifying annotations
            encryptionDetails.Permissions = PdfPermissions.ModifyAnnotations | PdfPermissions.DocumentAssembly;
            saveOptions.EncryptionDetails = encryptionDetails;

            // Render the document to PDF format with the specified permissions
            doc.Save(ArtifactsDir + "Rendering.EncryptionPermissions.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void SetNumeralFormat()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.NumeralFormat
            //ExFor:NumeralFormat
            //ExSummary:Demonstrates how to set the numeral format used when saving to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

            PdfSaveOptions options = new PdfSaveOptions();
            options.NumeralFormat = NumeralFormat.EasternArabicIndic;

            doc.Save(ArtifactsDir + "Rendering.SetNumeralFormat.pdf", options);
            //ExEnd
        }
    }
}