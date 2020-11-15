// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Imaging;
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

        [Test]
        public void OnePage()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions
            //ExFor:ImageSaveOptions.PageIndex
            //ExSummary:Shows how to render one page from a document to a JPEG image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            // Set the "PageIndex" to "1" to select the second page via
            // the zero-based index to start rendering the document from.
            options.PageIndex = 1;

            // When we save the document to the JPEG format, Aspose.Words only renders one page.
            // This image will contain one page starting from page two,
            // which will just be the second page of the original document.
            doc.Save(ArtifactsDir + "Rendering.OnePage.jpg", options);
            //ExEnd

            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "Rendering.OnePage.jpg");
        }

        [Test, Category("SkipMono")]
        public void PageByPage()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions
            //ExFor:ImageSaveOptions.PageIndex
            //ExFor:ImageSaveOptions.PageCount
            //ExSummary:Shows how to render every page of a document to a separate TIFF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

            // Set the "PageCount" property to "1" to render only one page of the document.
            // Many other image formats only render one page at a time, and do not use this property.
            options.PageCount = 1;

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Set the "PageIndex" property to the number of the first page from
                // which to start rendering the document from.
                options.PageIndex = i;

                doc.Save(ArtifactsDir + $"Rendering.PageByPage.{i + 1}.tiff", options);
            }
            //ExEnd

            List<string> imageFileNames = Directory.GetFiles(ArtifactsDir, "*.tiff")
                .Where(item => item.Contains("Rendering.PageByPage.") && item.EndsWith(".tiff")).ToList();

            Assert.AreEqual(3, imageFileNames.Count);

            foreach (string imageFileName in imageFileNames)
                TestUtil.VerifyImage(816, 1056, imageFileName);
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
            doc.Save(ArtifactsDir + "Rendering.SaveAsXpsBookFold.xps", xpsOptions);
            //ExEnd
        }

        [Test]
        public void SaveAsImage()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions.JpegQuality
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.#ctor
            //ExFor:ImageSaveOptions.JpegQuality
            //ExSummary:Shows how to configure compression while saving a document as a JPEG.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Logo.jpg");
            
            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
            // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
            imageOptions.JpegQuality = 10;

            doc.Save(ArtifactsDir + "Rendering.SaveAsImage.HighCompression.jpg", imageOptions);

            Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "Rendering.SaveAsImage.HighCompression.jpg").Length));

            // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
            // This will improve the quality of the image, but will also increase the file size.
            imageOptions.JpegQuality = 100;

            doc.Save(ArtifactsDir + "Rendering.SaveAsImage.HighQuality.jpg", imageOptions);

            Assert.That(60000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.SaveAsImage.HighQuality.jpg").Length));
            //ExEnd
        }

        [Test, Category("SkipMono")]
        public void SaveToTiffDefault()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Rendering.SaveToTiffDefault.tiff");
        }

        [TestCase(TiffCompression.None), Category("SkipMono")]
        [TestCase(TiffCompression.Rle), Category("SkipMono")]
        [TestCase(TiffCompression.Lzw), Category("SkipMono")]
        [TestCase(TiffCompression.Ccitt3), Category("SkipMono")]
        [TestCase(TiffCompression.Ccitt4), Category("SkipMono")]
        public void SaveToTiffCompression(TiffCompression tiffCompression)
        {
            //ExStart
            //ExFor:TiffCompression
            //ExFor:ImageSaveOptions.TiffCompression
            //ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

            // Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
            // which may result in a very large output file.
            // Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
            // Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
            // Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
            // Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
            options.TiffCompression = tiffCompression;

            doc.Save(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff", options);

            switch (tiffCompression)
            {
                case TiffCompression.None:
                    Assert.That(3000000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff").Length));
                    break;
                case TiffCompression.Rle:
                    Assert.That(600000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff").Length));
                    break;
                case TiffCompression.Lzw:
                    Assert.That(200000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff").Length));
                    break;
                case TiffCompression.Ccitt3:
                    Assert.That(90000, Is.AtLeast(new FileInfo(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff").Length));
                    break;
                case TiffCompression.Ccitt4:
                    Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "Rendering.SaveToTiffCompression.tiff").Length));
                    break;
            }
            //ExEnd
        }

        [Test]
        public void SetImageResolution()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.Resolution
            //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Set the "Resolution" property to "72" to render the document in 72dpi.
            options.Resolution = 72;
            
            doc.Save(ArtifactsDir + "Rendering.SetImageResolution.72dpi.png", options);

            Assert.That(120000, Is.AtLeast(new FileInfo(ArtifactsDir + "Rendering.SetImageResolution.72dpi.png").Length));

#if NET462 || JAVA
            Image image = Image.FromFile(ArtifactsDir + "Rendering.SetImageResolution.72dpi.png");

            Assert.AreEqual(612, image.Width);
            Assert.AreEqual(792, image.Height);
#elif NETCOREAPP2_1 || __MOBILE__
            using (SKBitmap image = SKBitmap.Decode(ArtifactsDir + "Rendering.SetImageResolution.72dpi.png")) 
            {
                Assert.AreEqual(612, image.Width);
                Assert.AreEqual(792, image.Height);
            }
#endif
            // Set the "Resolution" property to "300" to render the document in 300dpi.
            options.Resolution = 300;

            doc.Save(ArtifactsDir + "Rendering.SetImageResolution.300dpi.png", options);

            Assert.That(1100000, Is.LessThan(new FileInfo(ArtifactsDir + "Rendering.SetImageResolution.300dpi.png").Length));

#if NET462 || JAVA
            image = Image.FromFile(ArtifactsDir + "Rendering.SetImageResolution.300dpi.png");

            Assert.AreEqual(2550, image.Width);
            Assert.AreEqual(3300, image.Height);
#elif NETCOREAPP2_1 || __MOBILE__
            using (SKBitmap image = SKBitmap.Decode(ArtifactsDir + "Rendering.SetImageResolution.300dpi.png")) 
            {
                Assert.AreEqual(2550, image.Width);
                Assert.AreEqual(3300, image.Height);
            }
#endif
            //ExEnd
        }

        [Test]
        public void SetImagePaperColor()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.PaperColor
            //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);

            // Set the "PaperColor" property to a transparent color to apply a transparent
            // background to the document while rendering it to an image.
            imgOptions.PaperColor = Color.Transparent;

            doc.Save(ArtifactsDir + "Rendering.SetImagePaperColor.Transparent.png", imgOptions);

            // Set the "PaperColor" property to an opaque color to apply that color
            // as the background of the document as we render it to an image.
            imgOptions.PaperColor = Color.LightCoral;

            doc.Save(ArtifactsDir + "Rendering.SetImagePaperColor.LightCoral.png", imgOptions);
            //ExEnd

            TestUtil.ImageContainsTransparency(ArtifactsDir + "Rendering.SetImagePaperColor.Transparent.png");
            Assert.Throws<AssertionException>(() =>
                TestUtil.ImageContainsTransparency(ArtifactsDir + "Rendering.SetImagePaperColor.LightCoral.png"));
        }

        [Test]
        public void SaveToImageStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Save the document to a stream.
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Bmp);

                stream.Position = 0;
                
                // Read the stream back into an image.
#if NET462 || JAVA
                using (Image image = Image.FromStream(stream))
                {
                    Assert.AreEqual(ImageFormat.Bmp, image.RawFormat);
                    Assert.AreEqual(816, image.Width);
                    Assert.AreEqual(1056, image.Height);
                }
#elif NETCOREAPP2_1 || __MOBILE__
                using (SKBitmap image = SKBitmap.Decode(stream))
                {
                    Assert.AreEqual(816, image.Width);
                    Assert.AreEqual(1056, image.Height);
                }

                stream.Position = 0;

                SKCodec codec = SKCodec.Create(stream);

                Assert.AreEqual(SKEncodedImageFormat.Bmp, codec.EncodedFormat);
#endif
            }
            //ExEnd
        }

#if NET462 || JAVA
        [Test]
        public void RenderToSize()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Shows how to render a document to a bitmap at a specified location and size.
            Document doc = new Document(MyDir + "Rendering.docx");
            
            using (Bitmap bmp = new Bitmap(700, 700))
            {
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Set the "PageUnit" property to "GraphicsUnit.Inch" to use inches as the
                    // measurement unit for any transformations and dimensions that we will define.
                    gr.PageUnit = GraphicsUnit.Inch;

                    // Offset the output 0.5" from the edge.
                    gr.TranslateTransform(0.5f, 0.5f);

                    // Rotate the output by 10 degrees.
                    gr.RotateTransform(10);

                    // Draw a 3"x3" rectangle.
                    gr.DrawRectangle(new Pen(Color.Black, 3f / 72f), 0f, 0f, 3f, 3f);
                    
                    // Draw the first page of our document with the same dimensions and transformation as the rectangle.
                    // The rectangle will frame the first page.
                    float returnedScale = doc.RenderToSize(0, gr, 0f, 0f, 3f, 3f);

                    // This is the scaling factor that the RenderToSize method applied to the first page to fit the size we specified.
                    Assert.AreEqual(0.2566f, returnedScale, 0.0001f);

                    // Set the "PageUnit" property to "GraphicsUnit.Millimeter" to use millimeters as the
                    // measurement unit for any transformations and dimensions that we will define.
                    gr.PageUnit = GraphicsUnit.Millimeter;

                    // Reset the transformations that we used from the previous rendering.
                    gr.ResetTransform();

                    // Apply another set of transformations. 
                    gr.TranslateTransform(10, 10);
                    gr.ScaleTransform(0.5f, 0.5f);
                    gr.PageScale = 2f;

                    // Create another rectangle, and use it to frame another page from the document.
                    gr.DrawRectangle(new Pen(Color.Black, 1), 90, 10, 50, 100);
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
            //ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
            // The user opens or builds a document
            Document doc = new Document(MyDir + "Rendering.docx");

            // Calculate the number of rows and columns that we will fill with thumbnails.
            const int thumbColumns = 2;
            int thumbRows = Math.DivRem(doc.PageCount, thumbColumns, out int remainder);

            if (remainder > 0)
                thumbRows++;

            // Scale the thumbnails relative to the size of the first page. 
            const float scale = 0.25f;
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails.
            int imgWidth = thumbSize.Width * thumbColumns;
            int imgHeight = thumbSize.Height * thumbRows;
            
            using (Bitmap img = new Bitmap(imgWidth, imgHeight))
            {
                using (Graphics gr = Graphics.FromImage(img))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Fill the background, which is transparent by default, in white.
                    gr.FillRectangle(new SolidBrush(Color.White), 0, 0, imgWidth, imgHeight);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int rowIdx = Math.DivRem(pageIndex, thumbColumns, out int columnIdx);

                        // Specify where we want the thumbnail to appear.
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                        SizeF size = doc.RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);
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
            //ExSummary:Shows how to customize printing of Aspose.Words documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            MyPrintDocument printDoc = new MyPrintDocument(doc);
            printDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printDoc.PrinterSettings.FromPage = 1;
            printDoc.PrinterSettings.ToPage = 1;

            printDoc.Print();
        }

        /// <summary>
        /// Selects an appropriate paper size, orientation, and paper tray when printing.
        /// </summary>
        public class MyPrintDocument : PrintDocument
        {
            public MyPrintDocument(Document document)
            {
                mDocument = document;
            }

            /// <summary>
            /// Initializes the range of pages to be printed according to the user selection.
            /// </summary>
            protected override void OnBeginPrint(PrintEventArgs e)
            {
                base.OnBeginPrint(e);

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

                // A single Microsoft Word document can have multiple sections that specify pages with different sizes, 
                // orientations, and paper trays. The .NET printing framework calls this code before 
                // each page is printed, which gives us a chance to specify how to print the current page.
                PageInfo pageInfo = mDocument.GetPageInfo(mCurrentPage - 1);
                e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(PrinterSettings.PaperSizes);

                // Microsoft Word stores the paper source (printer tray) for each section as a printer-specific value.
                // To obtain the correct tray value you will need to use the RawKindValue, which your printer should return.
                e.PageSettings.PaperSource.RawKind = pageInfo.PaperTray;
                e.PageSettings.Landscape = pageInfo.Landscape;
            }

            /// <summary>
            /// Called for each page to render it for printing. 
            /// </summary>
            protected override void OnPrintPage(PrintPageEventArgs e)
            {
                base.OnPrintPage(e);

                // Aspose.Words rendering engine creates a page that is drawn from the origin (x = 0, y = 0) of the paper.
                // There will be a hard margin in the printer, which will render each page. We need to offset by that hard margin.
                float hardOffsetX, hardOffsetY;

                // Below are two ways of setting a hard margin.
                if (e.PageSettings != null && e.PageSettings.HardMarginX != 0 && e.PageSettings.HardMarginY != 0)
                {
                    // 1 -  Via the "PageSettings" property.
                    hardOffsetX = e.PageSettings.HardMarginX;
                    hardOffsetY = e.PageSettings.HardMarginY;
                }
                else
                {
                    // 2 -  Using our own values, if the "PageSettings" property is unavailable.
                    hardOffsetX = 20;
                    hardOffsetY = 20;
                }

                mDocument.RenderToScale(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f);

                mCurrentPage++;
                e.HasMorePages = mCurrentPage <= mPageTo;
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

            // The first section has 2 pages. We will assign a different printer paper tray to each one,
            // whose number will match a kind of paper source. These sources and their Kinds will vary
            // depending on the installed printer driver.
            PrinterSettings.PaperSourceCollection paperSources = new PrinterSettings().PaperSources;

            doc.FirstSection.PageSetup.FirstPageTray = paperSources[0].RawKind;
            doc.FirstSection.PageSetup.OtherPagesTray = paperSources[1].RawKind;

            Console.WriteLine("Document \"{0}\" contains {1} pages.", doc.OriginalFileName, doc.PageCount);

            float scale = 1.0f;
            float dpi = 96;

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Each page has a PageInfo object, whose index is the respective page's number.
                PageInfo pageInfo = doc.GetPageInfo(i);

                // Print the page's orientation and dimensions.
                Console.WriteLine($"Page {i + 1}:");
                Console.WriteLine($"\tOrientation:\t{(pageInfo.Landscape ? "Landscape" : "Portrait")}");
                Console.WriteLine($"\tPaper size:\t\t{pageInfo.PaperSize} ({pageInfo.WidthInPoints:F0}x{pageInfo.HeightInPoints:F0}pt)");
                Console.WriteLine($"\tSize in points:\t{pageInfo.SizeInPoints}");
                Console.WriteLine($"\tSize in pixels:\t{pageInfo.GetSizeInPixels(1.0f, 96)} at {scale * 100}% scale, {dpi} dpi");

                // Print the source tray information.
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
            // The "PrinterSettingsContainer" contains a "PrinterSettings" object,
            // which contains unique data for different printer drivers.
            PrinterSettingsContainer container = new PrinterSettingsContainer(new PrinterSettings());

            Console.WriteLine($"This printer contains {container.PaperSources.Count} printer paper sources:");
            foreach (PaperSource paperSource in container.PaperSources)
            {
                bool isDefault = container.DefaultPageSettingsPaperSource.SourceName == paperSource.SourceName;
                Console.WriteLine($"\t{paperSource.SourceName}, " +
                                  $"RawKind: {paperSource.RawKind} {(isDefault ? "(Default)" : "")}");
            }

            // The "PaperSizes" property contains the list of paper sizes that we can instruct the printer to use.
            // Both the PrinterSource and PrinterSize contain a "RawKind" attribute,
            // which equates to a paper type listed on the PaperSourceKind enum.
            // If there is a paper source with the same "RawKind" value as that of the page we are printing,
            // the printer will print the page using the provided paper source and size.
            // Otherwise, the printer will default to the source designated by the "DefaultPageSettingsPaperSource" property.
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
            //ExFor:Document.Print(String)
            //ExSummary:Shows how to print a document using the default printer.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Below are two ways of printing our document.
            // 1 -  Print using the default printer:
            doc.Print();

            // 2 -  Specify a printer that we wish to print the document with by name:
            string myPrinter = System.Drawing.Printing.PrinterSettings.InstalledPrinters[4];

            Assert.AreEqual("HPDAAB96 (HP ENVY 5000 series)", myPrinter);

            doc.Print(myPrinter);
            //ExEnd
        }
        
        [Ignore("Run only when the printer driver is installed")]
        [Test]
        public void PrintRange()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings)
            //ExFor:Document.Print(PrinterSettings, String)
            //ExSummary:Shows how to print a range of pages.
            Document doc = new Document(MyDir + "Rendering.docx");
            
            // Create a "PrinterSettings" object to modify the way in which we print the document.
            PrinterSettings printerSettings = new PrinterSettings();

            // Set the "PrintRange" property to "PrintRange.SomePages" to
            // tell the printer that we intend to print only some pages of the document.
            printerSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;

            // Set the "FromPage" property to "1", and the "ToPage" property to "3" to print pages 1 through to 3.
            // Page indexing is 1-based.
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            // Below are two ways of printing our document.
            // 1 -  Print while applying our printing settings:
            doc.Print(printerSettings);

            // 2 -  Print while applying our printing settings, while also
            // giving the document a custom name that we may recognize in the printer queue:
            doc.Print(printerSettings, "My rendered document");
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