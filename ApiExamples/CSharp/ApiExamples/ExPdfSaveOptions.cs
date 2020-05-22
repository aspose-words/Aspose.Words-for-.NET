// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
#if NET462 || NETCOREAPP2_1
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
        public void CreateMissingOutlineLevels()
        {
            //ExStart
            //ExFor:OutlineOptions.CreateMissingOutlineLevels
            //ExFor:ParagraphFormat.IsHeading
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExFor:PdfSaveOptions.SaveFormat
            //ExSummary:Shows how to create PDF document outline entries for headings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.1.1.1");
            builder.Writeln("Heading 1.1.1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9;

            builder.Writeln("Heading 1.1.1.1.1.1.1.1.1");
            builder.Writeln("Heading 1.1.1.1.1.1.1.1.2");

            // Create "PdfSaveOptions" with some mandatory parameters
            // "HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
            // "CreateMissingOutlineLevels" determining whether or not to create missing heading levels
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            // Bind PDF with Aspose.PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(11, bookmarks.Count);
            #endif
        }

        [Test]
        public void TableHeadingOutlines()
        {
            //ExStart
            //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
            //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
            // Create a blank document and insert a table with a heading-style text inside it
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Heading 1");
            builder.EndRow();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Write("Cell 1");
            builder.EndTable();

            // Create a PdfSaveOptions object that, when saving to .pdf with it, creates entries in the document outline for all headings levels 1-9,
            // and make sure headings inside tables are registered by the outline also
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateOutlinesForHeadingsInTables = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf");

            Assert.AreEqual(1, pdfDoc.Outlines.Count);
            Assert.AreEqual("Heading 1", pdfDoc.Outlines[1].Title);

            TableAbsorber tableAbsorber = new TableAbsorber();
            tableAbsorber.Visit(pdfDoc.Pages[1]);

            Assert.AreEqual("Heading 1", tableAbsorber.TableList[0].RowList[0].CellList[0].TextFragments[1].Text);
            Assert.AreEqual("Cell 1", tableAbsorber.TableList[0].RowList[1].CellList[0].TextFragments[1].Text);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void WithoutUpdateFields()
        {
            //ExStart
            //ExFor:PdfSaveOptions.Clone
            //ExFor:SaveOptions.UpdateFields
            //ExSummary:Shows how to update fields before saving into a PDF document.
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                UpdateFields = false
            };

            // PdfSaveOptions objects can be cloned
            Assert.AreNotSame(pdfSaveOptions, pdfSaveOptions.Clone());

            doc.Save(ArtifactsDir + "PdfSaveOptions.WithoutUpdateFields.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.WithoutUpdateFields.pdf");

            // Get text fragment by search String
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page  of");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are not updated
            Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments[1].Text);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void WithUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { UpdateFields = true };

            doc.Save(ArtifactsDir + "PdfSaveOptions.WithUpdateFields.pdf", pdfSaveOptions);

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.WithUpdateFields.pdf");

            // Get text fragment by search String from PDF document
            Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
            #endif
        }

        [Test]
        public void ImageCompression()
        {
            //ExStart
            //ExFor:PdfSaveOptions.Compliance
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExFor:PdfCompliance
            //ExFor:PdfImageColorSpaceExportMode
            //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
            Document doc = new Document(MyDir + "Images.docx");
            
            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,
                JpegQuality = 100, // Use JPEG compression at 50% quality to reduce file size
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf", optionsA1B);

            PdfSaveOptions optionsA1A = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,
                ExportDocumentStructure = true,
                ImageCompression = PdfImageCompression.Jpeg
            };

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf", optionsA1A);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            TestUtil.VerifyImage(2467, 1500, pdfDocImage.ToStream());
            
            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageCompression.PDF_A_1_B.pdf");
            pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.Throws<ArgumentException>(() => TestUtil.VerifyImage(2467, 1500, pdfDocImage.ToStream()));

            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageCompression.PDF_A_1_A.pdf");
            pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            TestUtil.VerifyImage(2467, 1500, pdfDocImage.ToStream());
            #endif
        }

        [Test]
        public void ColorRendering()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:ColorMode
            //ExFor:FixedPageSaveOptions.ColorMode
            //ExSummary:Shows how change image color with save options property
            Document doc = new Document(MyDir + "Images.docx");

            // Configure PdfSaveOptions to save every image in the input document in greyscale during conversion
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ColorMode = ColorMode.Grayscale };
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.AreEqual(1506, pdfDocImage.Width);
            Assert.AreEqual(918, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Grayscale, pdfDocImage.GetColorType());
            #endif
        }

        [Test]
        public void WindowsBarPdfTitle()
        {
            //ExStart
            //ExFor:PdfSaveOptions.DisplayDocTitle
            //ExSummary:Shows how to display title of the document as title bar.
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.BuiltInDocumentProperties.Title = "Windows bar pdf title";
            
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = true };

            doc.Save(ArtifactsDir + "PdfSaveOptions.WindowsBarPdfTitle.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.WindowsBarPdfTitle.pdf");

            Assert.IsTrue(pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            #endif
        }

        [Test]
        public void MemoryOptimization()
        {
            //ExStart
            //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
            //ExFor:SaveOptions.MemoryOptimization
            //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When set to true it will improve document memory footprint but will add extra time to processing
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            saveOptions.MemoryOptimization = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        [Test]
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

            // Set this property to false if you are sure that hyperlinks in document's model are already escaped
            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = isEscaped;
            options.OpenHyperlinksInNewWindow = true;

            builder.Document.Save(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf");

            // Get first page
            Page page = pdfDocument.Pages[1];
            // Get the first link annotation
            LinkAnnotation linkAnnot = (LinkAnnotation)page.Annotations[1];

            JavascriptAction action = (JavascriptAction)linkAnnot.Action;
            string uriText = action.Script;

            Assert.AreEqual(result, uriText);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void HandleBinaryRasterWarnings()
        {
            //ExStart
            //ExFor:MetafileRenderingMode
            //ExFor:MetafileRenderingOptions
            //ExFor:MetafileRenderingOptions.EmulateRasterOperations
            //ExFor:MetafileRenderingOptions.RenderingMode
            //ExFor:IWarningCallback
            //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
            //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions =
                new MetafileRenderingOptions
                {
                    EmulateRasterOperations = false,
                    RenderingMode = MetafileRenderingMode.VectorWithFallback
                };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

            Assert.AreEqual(1, callback.Warnings.Count);
            Assert.AreEqual("'R2_XORPEN' binary raster operation is partly supported.", callback.Warnings[0].Description);
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
                // For now type of warnings about unsupported metafile records changed from
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

            #if NET462 || NETCOREAPP2_1
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
		
		[Test]
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

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);
            Rectangle textFragmentRectangle = textAbsorber.TextFragments[3].Rectangle;

            if (doScaleWmfFonts)
                Assert.AreEqual(1.589d, textFragmentRectangle.Width, 0.001d);
            else
                Assert.AreEqual(5.045d, textFragmentRectangle.Width, 0.001d);
            #endif
        }

        [Test]
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

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);

            SetGlyphsPositionShowText tjOperator = (SetGlyphsPositionShowText)textAbsorber.TextFragments[1].Page.Contents[96];

            if (applyAdditionalTextPositioning)
                Assert.AreEqual("[0 (s) 0 (e) 1 (g) 0 (m) 0 (e) 0 (n) 0 (t) 0 (s) 0 ( ) 1 (o) 0 (f) 0 ( ) 1 (t) 0 (e) 0 (x) 0 (t)] TJ", tjOperator.ToString());
            else
                Assert.AreEqual("[(se) 1 (gments ) 1 (of ) 1 (text)] TJ", tjOperator.ToString());
            #endif
        }

        [Test]
        [TestCase(false)]
        [TestCase(true)]
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

            // In order to make a booklet, we will need to print this document, stack the pages
            // in the order they come out of the printer and then fold down the middle
            doc.Save(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1
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

            #if NET462 || NETCOREAPP2_1
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf");
            GoToAction action = (GoToAction)pdfDocument.OpenAction;

            Assert.AreEqual(0.25d, (action.Destination as XYZExplicitDestination).Zoom);
            #endif
        }

        [Test]
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

        [Test]
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
                Assert.Throws<AssertionException>(() => TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect", ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf"));
            }
        }

        [Test]
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

            // Add a custom document property that doesn't use the name of some built in properties
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
                    Assert.Throws<AssertionException>(() => TestUtil.FileContainsString(doc.CustomDocumentProperties[0].Name, 
                        ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                    Assert.Throws<AssertionException>(() => TestUtil.FileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", 
                        ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                    break;
                case PdfCustomPropertiesExport.Standard:
                    TestUtil.FileContainsString(doc.CustomDocumentProperties[0].Name, ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
                case PdfCustomPropertiesExport.Metadata:
                    TestUtil.FileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>", ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
            }
        }

        [Test]
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

            #if NET462 || NETCOREAPP2_1
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

        [Test]
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
            //ExSkip

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

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void ExportDocumentStructure(bool doExportStructure)
        {
            //ExStart
            //ExFor:PdfSaveOptions.ExportDocumentStructure
            //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro
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
        [Test]
        [TestCase(false)]
        [TestCase(true)]
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
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void PreblendImagesNetStandard2(bool doPreblendImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Transparent background logo.png"))
            {
                builder.InsertImage(image);
            }

            // Create a PdfSaveOptions object and setting this flag may change the quality and size of the output .pdf
            // because of the way some images are rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.PreblendImages = doPreblendImages;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf", options);
            //ExEnd

            TestPreblendImages(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf", doPreblendImages);
        }
#endif

        private void TestPreblendImages(string outFileName, bool doPreblendImages)
        {
#if NET462 || NETCOREAPP2_1
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
#endif
        }

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

        [Test]
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

#if NET462 || NETCOREAPP2_1
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