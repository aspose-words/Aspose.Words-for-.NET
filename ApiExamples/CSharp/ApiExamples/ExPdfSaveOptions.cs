// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using ColorMode = Aspose.Words.Saving.ColorMode;
using Document = Aspose.Words.Document;
using IWarningCallback = Aspose.Words.IWarningCallback;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;
using SaveOptions = Aspose.Words.Saving.SaveOptions;
using WarningInfo = Aspose.Words.WarningInfo;
using WarningType = Aspose.Words.WarningType;
using Image =
#if NET462 || JAVA
System.Drawing.Image;
#elif NETCOREAPP2_1 || __MOBILE__
SkiaSharp.SKBitmap;
using SkiaSharp;
#endif
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Operators;
using Aspose.Pdf.Text;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExPdfSaveOptions : ApiExampleBase
    {
        [Test]
        public void HeadingsOutlineLevels()
        {
            //ExStart
            //ExFor:OutlineOptions.CreateMissingOutlineLevels
            //ExFor:ParagraphFormat.IsHeading
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExFor:PdfSaveOptions.SaveFormat
            //ExSummary:Shows how to limit the level of headings that will appear in the outline of a saved PDF document.
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

            builder.Writeln("Heading 1.1.1");
            builder.Writeln("Heading 1.1.2");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Pdf;

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
            // The last two headings we have inserted above will not appear.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 2;
            saveOptions.OutlineOptions.CreateMissingOutlineLevels = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HeadingsOutlineLevels.pdf", saveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "PdfSaveOptions.HeadingsOutlineLevels.pdf");

            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(3, bookmarks.Count);
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void CreateMissingOutlineLevels(bool createMissingOutlineLevels)
        {
            //ExStart
            //ExFor:OutlineOptions.CreateMissingOutlineLevels
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExSummary:Shows how to work with outline levels that do not any corresponding headings when saving a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings that can serve as TOC entries of levels 1 and 5.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;

            builder.Writeln("Heading 1.1.1.1.1");
            builder.Writeln("Heading 1.1.1.1.2");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 5;

            // This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4. 
            // The output PDF document will treat outline levels 2, 3, and 4 as "missing".
            // Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
            // leaving blank outline entries since there are no usable headings.
            // Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
            // and treat the outline level 5 headings as level 2.
            saveOptions.OutlineOptions.CreateMissingOutlineLevels = createMissingOutlineLevels;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            if (createMissingOutlineLevels)
                Assert.AreEqual(6, bookmarks.Count);
            else
                Assert.AreEqual(3, bookmarks.Count);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void TableHeadingOutlines(bool createOutlinesForHeadingsInTables)
        {
            //ExStart
            //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
            //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table with three rows. The first row,
            // whose text we will format in a heading-type style, will serve as the column header.
            builder.StartTable();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Customers");
            builder.EndRow();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Write("John Doe");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Jane Doe");
            builder.EndTable();

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "1" to get the outline
            // to only register headings that have heading levels that are no larger than 1.
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 1;

            // Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
            // such as the one we have created above from the outline.
            // Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
            // in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
            pdfSaveOptions.OutlineOptions.CreateOutlinesForHeadingsInTables = createOutlinesForHeadingsInTables;

            doc.Save(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf");

            if (createOutlinesForHeadingsInTables)
            {
                Assert.AreEqual(1, pdfDoc.Outlines.Count);
                Assert.AreEqual("Customers", pdfDoc.Outlines[1].Title);
            } else
                Assert.AreEqual(0, pdfDoc.Outlines.Count);

            TableAbsorber tableAbsorber = new TableAbsorber();
            tableAbsorber.Visit(pdfDoc.Pages[1]);

            Assert.AreEqual("Customers", tableAbsorber.TableList[0].RowList[0].CellList[0].TextFragments[1].Text);
            Assert.AreEqual("John Doe", tableAbsorber.TableList[0].RowList[1].CellList[0].TextFragments[1].Text);
            Assert.AreEqual("Jane Doe", tableAbsorber.TableList[0].RowList[2].CellList[0].TextFragments[1].Text);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UpdateFields(bool updateFields)
        {
            //ExStart
            //ExFor:PdfSaveOptions.Clone
            //ExFor:SaveOptions.UpdateFields
            //ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
            // We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
            // each time we need them to display accurate values.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Hello World!");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
            // This is the preferable option if we know that all our fields will be up to date before saving.
            // Set the "UpdateFields" property to "true" to iterate through all the fields in the document
            // and update them before we save it as a PDF. This will make sure that all the fields will display the most accurate values in the PDF.
            options.UpdateFields = updateFields;
            
            // We can clone PdfSaveOptions objects.
            Assert.AreNotSame(options, options.Clone());

            doc.Save(ArtifactsDir + "PdfSaveOptions.UpdateFields.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.UpdateFields.pdf");

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            if (updateFields)
                Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
            else
                Assert.AreEqual("Page  of ", textFragmentAbsorber.TextFragments[1].Text);
            #endif
        }

        [TestCase(PdfCompliance.PdfA1b)]
        [TestCase(PdfCompliance.Pdf17)]
        [TestCase(PdfCompliance.PdfA1a)]
        public void Compliance(PdfCompliance pdfCompliance)
        {
            //ExStart
            //ExFor:PdfSaveOptions.Compliance
            //ExFor:PdfCompliance
            //ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
            // which aims to preserve the visual appearance of the document as Aspose.Words converts it to PDF.
            // Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
            // Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
            // which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
            // This helps with making documents searchable, but may significantly increase the size of already large documents.
            saveOptions.Compliance = pdfCompliance;

            doc.Save(ArtifactsDir + "PdfSaveOptions.Compliance.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.Compliance.pdf");

            switch (pdfCompliance)
            {
                case PdfCompliance.Pdf17:
                    Assert.AreEqual(PdfFormat.v_1_7, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.7", pdfDocument.Version);
                    break;
                case PdfCompliance.PdfA1a:
                    Assert.AreEqual(PdfFormat.PDF_A_1A, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.4", pdfDocument.Version);
                    break;
                case PdfCompliance.PdfA1b:
                    Assert.AreEqual(PdfFormat.PDF_A_1B, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.4", pdfDocument.Version);
                    break;
            }
#endif
        }

        [TestCase(PdfImageCompression.Auto)]
        [TestCase(PdfImageCompression.Jpeg)]
        public void ImageCompression(PdfImageCompression pdfImageCompression)
        {
            //ExStart
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jpeg image:");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertParagraph();
            builder.Writeln("Png image:");
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
            // "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
            // Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
            // "ImageCompression" property to control the quality of all images that end up in the output PDF.
            pdfSaveOptions.ImageCompression = pdfImageCompression;

            // Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
            pdfSaveOptions.JpegQuality = 10;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf");
            Stream pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();

            using (pdfDocImageStream)
            {
                TestUtil.VerifyImage(400, 400, pdfDocImageStream);
            }

            pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[2].ToStream();

            using (pdfDocImageStream)
            {
                switch (pdfImageCompression)
                {
                    case PdfImageCompression.Auto:
                        Assert.AreEqual(53700, new FileInfo(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf").Length, TestUtil.FileInfoLengthDelta);

                        Assert.Throws<ArgumentException>(() =>
                        {
                            TestUtil.VerifyImage(400, 400, pdfDocImageStream);
                        });
                        break;
                    case PdfImageCompression.Jpeg:
                        Assert.AreEqual(40000, new FileInfo(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf").Length, TestUtil.FileInfoLengthDelta);

                        TestUtil.VerifyImage(400, 400, pdfDocImageStream);
                        break;
                }
            }
#endif
        }

        [TestCase(PdfImageColorSpaceExportMode.Auto)]
        [TestCase(PdfImageColorSpaceExportMode.SimpleCmyk)]
        public void ImageColorSpaceExportMode(PdfImageColorSpaceExportMode pdfImageColorSpaceExportMode)
        {
            //ExStart
            //ExFor:PdfImageColorSpaceExportMode
            //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
            //ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jpeg image:");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertParagraph();
            builder.Writeln("Png image:");
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.Words to
            // automatically select the color space for images in the document that it converts to PDF. In most cases, the color space will be RGB.
            // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk" to use the CMYK color space for all images
            // in the saved PDF. Aspose.Words will also apply Flate compression to all images and ignore the value of the "ImageCompression" property.
            pdfSaveOptions.ImageColorSpaceExportMode = pdfImageColorSpaceExportMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageColorSpaceExportMode.pdf", pdfSaveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageColorSpaceExportMode.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            switch (pdfImageColorSpaceExportMode)
            {
                case PdfImageColorSpaceExportMode.Auto:
                    Assert.AreEqual(20115, pdfDocImage.ToStream().Length);
                    break;
                case PdfImageColorSpaceExportMode.SimpleCmyk:
                    Assert.AreEqual(138927, pdfDocImage.ToStream().Length);
                    break;
            }

            Assert.AreEqual(400, pdfDocImage.Width);
            Assert.AreEqual(400, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());

            pdfDocImage = pdfDocument.Pages[1].Resources.Images[2];

            switch (pdfImageColorSpaceExportMode)
            {
                case PdfImageColorSpaceExportMode.Auto:
                    Assert.AreEqual(19289, pdfDocImage.ToStream().Length);
                    break;
                case PdfImageColorSpaceExportMode.SimpleCmyk:
                    Assert.AreEqual(19980, pdfDocImage.ToStream().Length);
                    break;
            }

            Assert.AreEqual(400, pdfDocImage.Width);
            Assert.AreEqual(400, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
#endif
        }

        [Test]
        public void DownsampleOptions()
        {
            //ExStart
            //ExFor:DownsampleOptions
            //ExFor:DownsampleOptions.DownsampleImages
            //ExFor:DownsampleOptions.Resolution
            //ExFor:DownsampleOptions.ResolutionThreshold
            //ExFor:PdfSaveOptions.DownsampleOptions
            //ExSummary:Shows how to change the resolution of images in the PDF document.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // By default, Aspose.Words downsamples all images in a document that we save to PDF to 220 ppi.
            Assert.True(options.DownsampleOptions.DownsampleImages);
            Assert.AreEqual(220, options.DownsampleOptions.Resolution);
            Assert.AreEqual(0, options.DownsampleOptions.ResolutionThreshold);

            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.Default.pdf", options);

            // Set the "Resolution" property to "36" to downsample all images to 36 ppi.
            options.DownsampleOptions.Resolution = 36;

            // Set the "ResolutionThreshold" property to only apply the downsampling to
            // images with a resolution that is above 128 ppi.
            options.DownsampleOptions.ResolutionThreshold = 128;

            // Only the first two images from the document will be downsampled at this stage.
            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.LowerResolution.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.Default.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.AreEqual(399039, pdfDocImage.ToStream().Length);
            Assert.AreEqual(2467, pdfDocImage.Width);
            Assert.AreEqual(1500, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
#endif
        }

        [TestCase(ColorMode.Grayscale)]
        [TestCase(ColorMode.Normal)]
        public void ColorRendering(ColorMode colorMode)
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:ColorMode
            //ExFor:FixedPageSaveOptions.ColorMode
            //ExSummary:Shows how change image color with save options property.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            // Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
            // The size of the output document may be larger with this setting.
            // Set the "ColorMode" property to "Normal" to render all images in color.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ColorMode = colorMode };
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            switch (colorMode)
            {
                case ColorMode.Normal:
                    Assert.AreEqual(399039, pdfDocImage.ToStream().Length);
                    Assert.AreEqual(2467, pdfDocImage.Width);
                    Assert.AreEqual(1500, pdfDocImage.Height);
                    Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
                    break;
                case ColorMode.Grayscale:
                    Assert.AreEqual(1419611, pdfDocImage.ToStream().Length);
                    Assert.AreEqual(1506, pdfDocImage.Width);
                    Assert.AreEqual(918, pdfDocImage.Height);
                    Assert.AreEqual(ColorType.Grayscale, pdfDocImage.GetColorType());
                    break;
            }
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DocTitle(bool displayDocTitle)
        {
            //ExStart
            //ExFor:PdfSaveOptions.DisplayDocTitle
            //ExSummary:Shows how to display title of the document as title bar.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.BuiltInDocumentProperties.Title = "Windows bar pdf title";

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            // Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
            // to display the value of the document's "Title" built-in property in the tab that belongs to this document.
            // Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = displayDocTitle };

            doc.Save(ArtifactsDir + "PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DocTitle.pdf");

            Assert.AreEqual(displayDocTitle, pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void MemoryOptimization(bool memoryOptimization)
        {
            //ExStart
            //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
            //ExFor:SaveOptions.MemoryOptimization
            //ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

            // Set the "MemoryOptimization" property to "true" to lower the memory footprint of save operations
            // of large documents at the cost of increasing the duration of the operation.
            // Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
            saveOptions.MemoryOptimization = memoryOptimization;

            doc.Save(ArtifactsDir + "PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        [TestCase(@"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true)]
        [TestCase(@"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q= aspose\", true);", false)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", false)]
        public void EscapeUri(string uri, string result, bool isEscaped)
        {
            //ExStart
            //ExFor:PdfSaveOptions.EscapeUri
            //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
            //ExSummary:Shows how to escape hyperlinks in the document.
            DocumentBuilder builder = new DocumentBuilder();
            builder.InsertHyperlink("Testlink", uri, false);

            // Create a "PdfSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "EscapeUri" property to "true" if links in the document contain characters,
            // such as the blank space, that we need to replace with escape sequences, such as "%20".
            // Set the "EscapeUri" property to "false" if we are sure that the links
            // in this document no not need any such escape character substitution.
            options.EscapeUri = isEscaped;
            options.OpenHyperlinksInNewWindow = true;

            builder.Document.Save(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf");

            Page page = pdfDocument.Pages[1];
            LinkAnnotation linkAnnot = (LinkAnnotation)page.Annotations[1];

            JavascriptAction action = (JavascriptAction)linkAnnot.Action;
            string uriText = action.Script;

            Assert.AreEqual(result, uriText);
            #endif
        }

        //ExStart
        //ExFor:MetafileRenderingMode
        //ExFor:MetafileRenderingOptions
        //ExFor:MetafileRenderingOptions.EmulateRasterOperations
        //ExFor:MetafileRenderingOptions.RenderingMode
        //ExFor:IWarningCallback
        //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
        //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
        [Test, Category("SkipMono")] //ExSkip
        public void HandleBinaryRasterWarnings()
        {
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions =
                new MetafileRenderingOptions
                {
                    EmulateRasterOperations = false,
                    RenderingMode = MetafileRenderingMode.VectorWithFallback
                };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words
            // renders this metafile to a bitmap
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

            Assert.AreEqual(1, callback.Warnings.Count);
            Assert.AreEqual("'R2_XORPEN' binary raster operation is partly supported.",
                callback.Warnings[0].Description);
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
                // For now, type of warnings about unsupported metafile records changed from
                // DataLoss/UnexpectedContent to MinorFormattingLoss
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    Warnings.Warning(info);
                }
            }

            public WarningInfoCollection Warnings = new WarningInfoCollection();
        }
        //ExEnd

        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All)]
        public void HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
        {
            //ExStart
            //ExFor:HeaderFooterBookmarksExportMode
            //ExFor:OutlineOptions
            //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
            //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
            //ExFor:PdfSaveOptions.PageMode
            //ExFor:PdfPageMode
            //ExSummary:Shows how bookmarks in headers/footers are exported to pdf.
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            // You can specify how bookmarks in headers/footers are exported
            // There is a several options for this:
            // "None" - Bookmarks in headers/footers are not exported
            // "First" - Only bookmark in first header/footer of the section is exported
            // "All" - Bookmarks in all headers/footers are exported
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                HeaderFooterBookmarksExportMode = headerFooterBookmarksExportMode,
                OutlineOptions = { DefaultBookmarksOutlineLevel = 1 },
                PageMode = PdfPageMode.UseOutlines
            };
            doc.Save(ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");
            string inputDocLocaleName = new CultureInfo(doc.Styles.DefaultFont.LocaleId).Name;

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDoc.Pages.Accept(textFragmentAbsorber);
            switch (headerFooterBookmarksExportMode)
            {
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({inputDocLocaleName})>>\r\n", 
                        ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                    Assert.AreEqual(0, pdfDoc.Outlines.Count);
                    break;
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First:
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/Outlines 13 0 R/PageMode /UseOutlines/Lang({inputDocLocaleName})>>", 
                        ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                    OutlineCollection outlineItemCollection = pdfDoc.Outlines;

                    Assert.AreEqual(4, outlineItemCollection.Count);
                    Assert.AreEqual("Bookmark_1", outlineItemCollection[1].Title);
                    Assert.AreEqual("1 XYZ 233 806 0", outlineItemCollection[1].Destination.ToString());

                    Assert.AreEqual("Bookmark_2", outlineItemCollection[2].Title);
                    Assert.AreEqual("1 XYZ 84 47 0", outlineItemCollection[2].Destination.ToString());

                    Assert.AreEqual("Bookmark_3", outlineItemCollection[3].Title);
                    Assert.AreEqual("2 XYZ 85 806 0", outlineItemCollection[3].Destination.ToString());

                    Assert.AreEqual("Bookmark_4", outlineItemCollection[4].Title);
                    Assert.AreEqual("2 XYZ 85 48 0", outlineItemCollection[4].Destination.ToString());
                    break;
            }
            #endif
        }

        [Test]
        public void UnsupportedImageFormatWarning()
        {
            Document doc = new Document(MyDir + "Corrupted image.docx");

            SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
            doc.WarningCallback = saveWarningCallback;

            doc.Save(ArtifactsDir + "PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat.Pdf);

            Assert.That(saveWarningCallback.SaveWarnings[0].Description,
                Is.EqualTo("Image can not be processed. Possibly unsupported image format."));
        }

        public class SaveWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine($"{info.WarningType}: {info.Description}.");
                    SaveWarnings.Warning(info);
                }
            }

            internal WarningInfoCollection SaveWarnings = new WarningInfoCollection();
		}
		
		[TestCase(false)]
        [TestCase(true)]
        public void FontsScaledToMetafileSize(bool doScaleWmfFonts)
        {
            //ExStart
            //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
            //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
            Document doc = new Document(MyDir + "WMF with text.docx");

            // There is a several options for this:
            // 'True' - Aspose.Words emulates font scaling according to metafile size on the page
            // 'False' - Aspose.Words displays the fonts as metafile is rendered to its default size
            // Use 'False' option is used only when metafile is rendered as vector graphics
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions.ScaleWmfFontsToMetafileSize = doScaleWmfFonts;

            doc.Save(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);
            Rectangle textFragmentRectangle = textAbsorber.TextFragments[3].Rectangle;

            Assert.AreEqual(doScaleWmfFonts ? 1.589d : 5.045d, textFragmentRectangle.Width, 0.001d);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AdditionalTextPositioning(bool applyAdditionalTextPositioning)
        {
            //ExStart
            //ExFor:PdfSaveOptions.AdditionalTextPositioning
            //ExSummary:Show how to write additional text positioning operators.
            Document doc = new Document(MyDir + "Rendering.docx");

            // This may help to overcome issues with inaccurate text positioning with some printers, even if the PDF looks fine,
            // but the file size will increase due to higher text positioning precision used
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                AdditionalTextPositioning = applyAdditionalTextPositioning,
                TextCompression = PdfTextCompression.None
            };

            doc.Save(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);

            SetGlyphsPositionShowText tjOperator = (SetGlyphsPositionShowText)textAbsorber.TextFragments[1].Page.Contents[96];

            Assert.AreEqual(
                applyAdditionalTextPositioning
                    ? "[0 (s) 0 (e) 1 (g) 0 (m) 0 (e) 0 (n) 0 (t) 0 (s) 0 ( ) 1 (o) 0 (f) 0 ( ) 1 (t) 0 (e) 0 (x) 0 (t)] TJ"
                    : "[(se) 1 (gments ) 1 (of ) 1 (text)] TJ", tjOperator.ToString());
#endif
        }

        [TestCase(false, Category = "SkipMono")]
        [TestCase(true, Category = "SkipMono")]
        public void SaveAsPdfBookFold(bool doRenderTextAsBookfold)
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
            // Open a document with multiple paragraphs
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Configure both page setup and PdfSaveOptions to create a book fold
            foreach (Section s in doc.Sections)
            {
                s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
            }

            PdfSaveOptions options = new PdfSaveOptions();
            options.UseBookFoldPrintingSettings = doRenderTextAsBookfold;

            // Once we print this document, we can turn it into a booklet by stacking the pages
            // in the order they come out of the printer and then folding down the middle
            doc.Save(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();

            pdfDocument.Pages.Accept(textAbsorber);

            if (doRenderTextAsBookfold)
            {
                Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal));
                Assert.False(textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal));
                Assert.False(textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.Ordinal));
            }
            else
            {
                Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.Ordinal));
            }
            #endif
        }

        [Test]
        public void ZoomBehaviour()
        {
            //ExStart
            //ExFor:PdfSaveOptions.ZoomBehavior
            //ExFor:PdfSaveOptions.ZoomFactor
            //ExFor:PdfZoomBehavior
            //ExSummary:Shows how to set the default zooming of an output PDF to 1/4 of default size.
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions options = new PdfSaveOptions
            {
                ZoomBehavior = PdfZoomBehavior.ZoomFactor,
                ZoomFactor = 25,
            };

            // Upon opening the .pdf with a viewer such as Adobe Acrobat Pro, the zoom level will be at 25% by default,
            // with thumbnails for each page to the left
            doc.Save(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf");
            GoToAction action = (GoToAction)pdfDocument.OpenAction;

            Assert.AreEqual(0.25d, (action.Destination as XYZExplicitDestination).Zoom);
            #endif
        }

        [TestCase(PdfPageMode.FullScreen)]
        [TestCase(PdfPageMode.UseThumbs)]
        [TestCase(PdfPageMode.UseOC)]
        [TestCase(PdfPageMode.UseOutlines)]
        [TestCase(PdfPageMode.UseNone)]
        public void PageMode(PdfPageMode pageMode)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PageMode
            //ExFor:PdfPageMode
            //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
            Document doc = new Document(MyDir + "Document.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.PageMode = pageMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PageMode.pdf", options);
            //ExEnd
            
            string docLocaleName = new CultureInfo(doc.Styles.DefaultFont.LocaleId).Name;

            switch (pageMode)
            {
                case PdfPageMode.FullScreen:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({docLocaleName})>>\r\n", ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseThumbs:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({docLocaleName})>>", ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseOC:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({docLocaleName})>>\r\n", ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseOutlines:
                case PdfPageMode.UseNone:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({docLocaleName})>>\r\n", ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
            }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void NoteHyperlinks(bool doCreateHyperlinks)
        {
            //ExStart
            //ExFor:PdfSaveOptions.CreateNoteHyperlinks
            //ExSummary:Shows how to make footnotes and endnotes work like hyperlinks.
            // Open a document with footnotes/endnotes
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");

            // Creating a PdfSaveOptions instance with this flag set will convert footnote/endnote number symbols in the text
            // into hyperlinks pointing to the footnotes, and the actual footnotes/endnotes at the end of pages into links to their
            // referenced body text
            PdfSaveOptions options = new PdfSaveOptions();
            options.CreateNoteHyperlinks = doCreateHyperlinks;

            doc.Save(ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf", options);
            //ExEnd

            if (doCreateHyperlinks)
            {
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 677 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 79 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 654 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 68 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19905853 88.66500092 79.69805908]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 202 733 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70005798 88.66500092 68.19905853]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 258 711 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 157 733 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 212 711 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
            }
            else
            {
                if (!IsRunningOnMono())
                    Assert.Throws<AssertionException>(() => TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect", ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf"));
            }
        }

        [TestCase(PdfCustomPropertiesExport.None)]
        [TestCase(PdfCustomPropertiesExport.Standard)]
        [TestCase(PdfCustomPropertiesExport.Metadata)]
        public void CustomPropertiesExport(PdfCustomPropertiesExport pdfCustomPropertiesExportMode)
        {
            //ExStart
            //ExFor:PdfCustomPropertiesExport
            //ExFor:PdfSaveOptions.CustomPropertiesExport
            //ExSummary:Shows how to export custom properties while saving to .pdf.
            Document doc = new Document();

            // Add a custom document property that does not use the name of some built in properties
            doc.CustomDocumentProperties.Add("Company", "My value");
            
            // Configure the PdfSaveOptions like this will display the properties
            // in the "Document Properties" menu of Adobe Acrobat Pro
            PdfSaveOptions options = new PdfSaveOptions();
            options.CustomPropertiesExport = pdfCustomPropertiesExportMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            //ExEnd

            switch (pdfCustomPropertiesExportMode)
            {
                case PdfCustomPropertiesExport.None:
                    if (!IsRunningOnMono())
                    {
                        Assert.Throws<AssertionException>(() => TestUtil.FileContainsString(doc.CustomDocumentProperties[0].Name,
                            ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                        Assert.Throws<AssertionException>(() => TestUtil.FileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                            ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                    }
                    break;
                case PdfCustomPropertiesExport.Standard:
                    TestUtil.FileContainsString(doc.CustomDocumentProperties[0].Name, ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
                case PdfCustomPropertiesExport.Metadata:
                    TestUtil.FileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
            }
        }

        [TestCase(DmlEffectsRenderingMode.None)]
        [TestCase(DmlEffectsRenderingMode.Simplified)]
        [TestCase(DmlEffectsRenderingMode.Fine)]
        public void DrawingMLEffects(DmlEffectsRenderingMode effectsRenderingMode)
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:DmlEffectsRenderingMode
            //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
            //ExFor:SaveOptions.DmlEffectsRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to configure DrawingML rendering quality with PdfSaveOptions.
            Document doc = new Document(MyDir + "DrawingML shape effects.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.DmlEffectsRenderingMode = effectsRenderingMode;

            Assert.AreEqual(DmlRenderingMode.DrawingML, options.DmlRenderingMode);

            doc.Save(ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");

            ImagePlacementAbsorber imb = new ImagePlacementAbsorber();
            imb.Visit(pdfDocument.Pages[1]);

            TableAbsorber ttb = new TableAbsorber();
            ttb.Visit(pdfDocument.Pages[1]);

            switch (effectsRenderingMode)
            {
                case DmlEffectsRenderingMode.None:
                case DmlEffectsRenderingMode.Simplified:
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");
                    Assert.AreEqual(0, imb.ImagePlacements.Count);
                    Assert.AreEqual(28, ttb.TableList.Count);
                    break;
                case DmlEffectsRenderingMode.Fine:
                    TestUtil.FileContainsString("4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>/XObject<</X1 9 0 R/X2 10 0 R/X3 11 0 R/X4 12 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");
                    Assert.AreEqual(21, imb.ImagePlacements.Count);
                    Assert.AreEqual(4, ttb.TableList.Count);
                    break;
            }
            #endif
        }

        [TestCase(DmlRenderingMode.Fallback)]
        [TestCase(DmlRenderingMode.DrawingML)]
        public void DrawingMLFallback(DmlRenderingMode dmlRenderingMode)
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to render fallback shapes when saving to Pdf.
            Document doc = new Document(MyDir + "DrawingML shape fallbacks.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.DmlRenderingMode = dmlRenderingMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf", options);
            //ExEnd

            switch (dmlRenderingMode)
            {
                case DmlRenderingMode.DrawingML:
                    TestUtil.FileContainsString("<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
                    break;
                case DmlRenderingMode.Fallback:
                    TestUtil.FileContainsString("4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group ",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
                    break;
            }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportDocumentStructure(bool doExportStructure)
        {
            //ExStart
            //ExFor:PdfSaveOptions.ExportDocumentStructure
            //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased, and the structure will be visible in the "Content" navigation pane of Adobe Acrobat Pro
            PdfSaveOptions options = new PdfSaveOptions();
            options.ExportDocumentStructure = doExportStructure;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            //ExEnd

            if (doExportStructure)
            {
                TestUtil.FileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                    ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf");
            }
            else
            {
                TestUtil.FileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf");
            }
        }

#if NET462 || JAVA
        [TestCase(false, Category = "SkipMono")]
        [TestCase(true, Category = "SkipMono")]
        public void PreblendImages(bool doPreblendImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(ImageDir + "Transparent background logo.png");
            builder.InsertImage(img);

            // Setting this flag in a SaveOptions object may change the quality and size of the output .pdf
            // because of the way some images are rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.PreblendImages = doPreblendImages;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf", options);
            //ExEnd

            TestPreblendImages(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf", doPreblendImages);
        }

        private void TestPreblendImages(string outFileName, bool doPreblendImages)
        {
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(outFileName);
            XImage image = pdfDocument.Pages[1].Resources.Images[1];

            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream);

                if (doPreblendImages)
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20849 ", outFileName);
                    Assert.AreEqual(17898, stream.Length);
                }
                else
                {
                    TestUtil.FileContainsString("9 0 obj\r\n19289 ", outFileName);
                    Assert.AreEqual(19216, stream.Length);
                }
            }
        }

        [Test]
        public void InterpolateImages()
        {
            //ExStart
            //ExFor:PdfSaveOptions.InterpolateImages
            //ExSummary:Shows how to improve the quality of an image in the rendered documents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(ImageDir + "Transparent background logo.png");
            builder.InsertImage(img);
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
            //ExEnd
        }

        [Test, Category("SkipMono")]
        public void Dml3DEffectsRenderingModeTest()
        {
            Document doc = new Document(MyDir + "DrawingML shape 3D effects.docx");
            
            RenderCallback warningCallback = new RenderCallback();
            doc.WarningCallback = warningCallback;
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

            Assert.AreEqual(43, warningCallback.Count);
        }

        public class RenderCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                Console.WriteLine($"{info.WarningType}: {info.Description}.");
                mWarnings.Add(info);
            }

            public WarningInfo this[int i] => mWarnings[i];

            /// <summary>
            /// Clears warning collection.
            /// </summary>
            public void Clear()
            {
                mWarnings.Clear();
            }

            public int Count => mWarnings.Count;

            /// <summary>
            /// Returns true if a warning with the specified properties has been generated.
            /// </summary>
            public bool Contains(WarningSource source, WarningType type, string description)
            {
                return mWarnings.Any(warning => warning.Source == source && warning.WarningType == type && warning.Description == description);
            }

            private readonly List<WarningInfo> mWarnings = new List<WarningInfo>();
        }

#elif NETCOREAPP2_1
        [TestCase(false)]
        [TestCase(true)]
        public void PreblendImagesNetStandard2(bool doPreblendImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Image image = Image.Decode(ImageDir + "Transparent background logo.png"))
            {
                builder.InsertImage(image);
            }

            // Create a PdfSaveOptions object and setting this flag may change the quality and size of the output .pdf
            // because of the way some images are rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.PreblendImages = doPreblendImages;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf", options);
            //ExEnd

            TestPreblendImagesNetStandard2(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf", doPreblendImages);
        }

        private void TestPreblendImagesNetStandard2(string outFileName, bool doPreblendImages)
        {
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(outFileName);
            XImage image = pdfDocument.Pages[1].Resources.Images[1];

            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream);

                if (doPreblendImages)
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20849 ", outFileName);
                    Assert.AreEqual(17898, stream.Length);
                }
                else
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20266 ", outFileName);
                    Assert.AreEqual(19135, stream.Length);
                }
            }
        }

        [Test]
        public void InterpolateImages()
        {
            //ExStart
            //ExFor:PdfSaveOptions.InterpolateImages
            //ExSummary:Shows how to improve the quality of an image in the rendered documents (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Image image = Image.Decode(ImageDir + "Transparent background logo.png"))
            {
                builder.InsertImage(image);
            }
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
            //ExEnd
        }
#endif

        [Test]
        public void PdfDigitalSignature()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor
            //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
            //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
            //ExFor:PdfDigitalSignatureDetails.Location
            //ExFor:PdfDigitalSignatureDetails.Reason
            //ExFor:PdfDigitalSignatureDetails.SignatureDate
            //ExFor:PdfDigitalSignatureHashAlgorithm
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExSummary:Shows how to sign a generated PDF using Aspose.Words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Signed PDF contents.");

            // Load the certificate from disk
            // The other constructor overloads can be used to load certificates from different locations
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Pass the certificate and details to the save options class to sign with
            PdfSaveOptions options = new PdfSaveOptions();
            DateTime signingTime = DateTime.Now;
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", signingTime);

            // We can use this attribute to set a different hash algorithm
            options.DigitalSignatureDetails.HashAlgorithm = PdfDigitalSignatureHashAlgorithm.Sha256;

            Assert.AreEqual("Test Signing", options.DigitalSignatureDetails.Reason);
            Assert.AreEqual("Aspose Office", options.DigitalSignatureDetails.Location);
            Assert.AreEqual(signingTime.ToUniversalTime(), options.DigitalSignatureDetails.SignatureDate);

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
            //ExEnd
            
            TestUtil.FileContainsString("6 0 obj\r\n" +
                                        "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\0s\0p\0o\0s\0e\0D\0i\0g\0i\0t\0a\0l\0S\0i\0g\0n\0a\0t\0u\0r\0e)/AP <</N 8 0 R>>>>", 
                ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf");

            Assert.False(FileFormatUtil.DetectFileFormat(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf").HasDigitalSignature);
        }

        [Test]
        public void PdfDigitalSignatureTimestamp()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails.TimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
            //ExFor:PdfDigitalSignatureTimestampSettings.Password
            //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
            //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
            //ExFor:PdfDigitalSignatureTimestampSettings.UserName
            //ExSummary:Shows how to sign a generated PDF and timestamp it using Aspose.Words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Signed PDF contents.");

            // Create a digital signature for the document that we will save
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.Now);

            // We can set a verified timestamp for our signature as well, with a valid timestamp authority
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword");

            // The default lifespan of the timestamp is 100 seconds
            Assert.AreEqual(100.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);

            // We can set our own timeout period via the constructor
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", TimeSpan.FromMinutes(30));

            Assert.AreEqual(1800.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);
            Assert.AreEqual("https://freetsa.org/tsr", options.DigitalSignatureDetails.TimestampSettings.ServerUrl);
            Assert.AreEqual("JohnDoe", options.DigitalSignatureDetails.TimestampSettings.UserName);
            Assert.AreEqual("MyPassword", options.DigitalSignatureDetails.TimestampSettings.Password);

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
            //ExEnd

            Assert.False(FileFormatUtil.DetectFileFormat(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").HasDigitalSignature);
            TestUtil.FileContainsString("6 0 obj\r\n" +
                                        "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\0s\0p\0o\0s\0e\0D\0i\0g\0i\0t\0a\0l\0S\0i\0g\0n\0a\0t\0u\0r\0e)/AP <</N 8 0 R>>>>", 
            ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
        }

        [TestCase(EmfPlusDualRenderingMode.Emf)]
        [TestCase(EmfPlusDualRenderingMode.EmfPlus)]
        [TestCase(EmfPlusDualRenderingMode.EmfPlusWithFallback)]
        public void RenderMetafile(EmfPlusDualRenderingMode renderingMode)
        {
            //ExStart
            //ExFor:EmfPlusDualRenderingMode
            //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
            //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
            //ExSummary:Shows how to adjust EMF (Enhanced Windows Metafile) rendering options when saving to PDF.
            Document doc = new Document(MyDir + "EMF.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions.EmfPlusDualRenderingMode = renderingMode;
            saveOptions.MetafileRenderingOptions.UseEmfEmbeddedToWmf = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");

            switch (renderingMode)
            {
                case EmfPlusDualRenderingMode.Emf:
                case EmfPlusDualRenderingMode.EmfPlusWithFallback:
                    Assert.AreEqual(0, pdfDocument.Pages[1].Resources.Images.Count);
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R/FAAABD 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");
                    break;
                case EmfPlusDualRenderingMode.EmfPlus:
                    Assert.AreEqual(1, pdfDocument.Pages[1].Resources.Images.Count);
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>/XObject<</X1 9 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");
                    break;
            }
#endif
        }
    }
}