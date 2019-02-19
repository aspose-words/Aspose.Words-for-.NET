// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
using Document = Aspose.Words.Document;
using IWarningCallback = Aspose.Words.IWarningCallback;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;
using SaveOptions = Aspose.Words.Saving.SaveOptions;
using WarningInfo = Aspose.Words.WarningInfo;
using WarningType = Aspose.Words.WarningType;
#if !(__MOBILE__ || MAC)
using Aspose.Pdf.Facades;
using Aspose.Pdf.Annotations;
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
            //ExSummary:Shows how to create missing outline levels saving the document in PDF
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

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

            doc.Save(ArtifactsDir + "CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd
#if !(__MOBILE__ || MAC)
            // Bind PDF with Aspose.PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "CreateMissingOutlineLevels.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(11, bookmarks.Count);
#endif
        }

        [Test]
        public void AllowToAddBookmarksWithWhiteSpaces()
        {
            //ExStart
            //ExFor:OutlineOptions.BookmarksOutlineLevels
            //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
            //ExSummary:Shows how adding bookmarks outlines with whitespaces(pdf, xps)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add bookmarks with whitespaces. MS Word formats (like doc, docx) does not support bookmarks with whitespaces by default 
            // and all whitespaces in the bookmarks were replaced with underscores. If you need to use bookmarks in PDF or XPS outlines, you can use them with whitespaces.
            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");

            // Specify bookmarks outline level. If you are using xps format, just use XpsSaveOptions.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
            pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);

            doc.Save(ArtifactsDir + "Bookmarks.WhiteSpaces.pdf", pdfSaveOptions);
            //ExEnd
#if !(__MOBILE__ || MAC)
            // Bind pdf with Aspose.Pdf
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "Bookmarks.WhiteSpaces.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(2, bookmarks.Count);

            // Assert that all the bookmarks title are with whitespaces
            Assert.AreEqual("My Bookmark", bookmarks[0].Title);
            Assert.AreEqual("Nested Bookmark", bookmarks[1].Title);
#endif
        }

        //Note: Test doesn't contain validation result.
        //For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
        [Test]
        public void DrawingMl()
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to define rendering for DML shapes
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DmlRenderingMode = DmlRenderingMode.DrawingML };

            doc.Save(ArtifactsDir + "DrawingMl.pdf", pdfSaveOptions);
            //ExEnd
        }

        [Test]
        [Category("SkipMono")]
        public void WithoutUpdateFields()
        {
            //ExStart
            //ExFor:SaveOptions.UpdateFields
            //ExSummary:Shows how to update fields before saving into a PDF document.
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                UpdateFields = false
            };

            doc.Save(ArtifactsDir + "UpdateFields_False.pdf", pdfSaveOptions);
            //ExEnd
#if !(__MOBILE__ || MAC)
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "UpdateFields_False.pdf");

            // Get text fragment by search String
            Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page  of");
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

            doc.Save(ArtifactsDir + "UpdateFields_False.pdf", pdfSaveOptions);
#if !(__MOBILE__ || MAC)
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "UpdateFields_False.pdf");

            // Get text fragment by search String from PDF document
            Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
#endif
        }

        // For assert this test you need to open "SaveOptions.PdfImageCompression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageCompression PDF_A_1_A Out.pdf" 
        // and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
        [Test]
        public void ImageCompression()
        {
            //ExStart
            //ExFor:PdfSaveOptions.Compliance
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExFor:PdfCompliance
            //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
            Document doc = new Document(MyDir + "SaveOptions.PdfImageCompression.rtf");
            
            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            doc.Save(ArtifactsDir + "SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions();
            optionsA1B.Compliance = PdfCompliance.PdfA1b;
            optionsA1B.ImageCompression = PdfImageCompression.Jpeg;
            optionsA1B.JpegQuality = 100; // Use JPEG compression at 50% quality to reduce file size.

            doc.Save(ArtifactsDir + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
            //ExEnd

            PdfSaveOptions optionsA1A = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,
                ExportDocumentStructure = true,
                ImageCompression = PdfImageCompression.Jpeg
            };

            doc.Save(ArtifactsDir + "SaveOptions.PdfImageComppression PDF_A_1_A.pdf", optionsA1A);
        }

        [Test]
        public void ColorRendering()
        {
            //ExStart
            //ExFor:SaveOptions.ColorMode
            //ExSummary:Shows how change image color with save options property
            // Open document with color image
            Document doc = new Document(MyDir + "Rendering.doc");
            // Set grayscale mode for document
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ColorMode = ColorMode.Grayscale };

            // Assert that color image in document was grey
            doc.Save(ArtifactsDir + "ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
            //ExEnd
        }

        [Test]
        public void WindowsBarPdfTitle()
        {
            //ExStart
            //ExFor:PdfSaveOptions.DisplayDocTitle
            //ExSummary:Shows how to display title of the document as title bar.
            Document doc = new Document(MyDir + "Rendering.doc");
            doc.BuiltInDocumentProperties.Title = "Windows bar pdf title";
            
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = true };

            doc.Save(ArtifactsDir + "PdfTitle.pdf", pdfSaveOptions);
            //ExEnd
#if !(__MOBILE__ || MAC)
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfTitle.pdf");

            Assert.IsTrue(pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
#endif
        }

        [Test]
        public void MemoryOptimization()
        {
            //ExStart
            //ExFor:SaveOptions.MemoryOptimization
            //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
            Document doc = new Document(MyDir + "SaveOptions.MemoryOptimization.doc");
            // When set to true it will improve document memory footprint but will add extra time to processing. 
            // This optimization is only applied during save operation.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            saveOptions.MemoryOptimization = true;

            doc.Save(ArtifactsDir + "SaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        [TestCase(@"https://www.google.com/search?q= aspose", @"https://www.google.com/search?q=%20aspose", true)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", @"https://www.google.com/search?q=%20aspose", true)]
        [TestCase(@"https://www.google.com/search?q= aspose", @"https://www.google.com/search?q= aspose", false)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", @"https://www.google.com/search?q=%20aspose", false)]
        public void EscapeUri(string uri, string result, bool isEscaped)
        {
            //ExStart
            //ExFor:PdfSaveOptions.EscapeUri
            //ExSummary: Shows how to escape hyperlinks or not in the document.
            DocumentBuilder builder = new DocumentBuilder();
            builder.InsertHyperlink("Testlink", uri, false);

            // Set this property to false if you are sure that hyperlinks in document's model are already escaped
            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = isEscaped;

            builder.Document.Save(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf", options);
            //ExEnd

#if !(__MOBILE__ || MAC)
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf");

            // Get first page
            Aspose.Pdf.Page page = pdfDocument.Pages[1];
            // Get the first link annotation
            LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

            GoToURIAction action = (GoToURIAction) linkAnnot.Action;
            string uriText = action.URI;

            Assert.AreEqual(result, uriText);
#endif
            //ExEnd
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
            //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records
            Document doc = new Document(MyDir + "PdfSaveOptions.HandleRasterWarnings.doc");

            MetafileRenderingOptions metafileRenderingOptions =
                new MetafileRenderingOptions
                {
                    EmulateRasterOperations = false,
                    RenderingMode = MetafileRenderingMode.VectorWithFallback
                };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.HandleRasterWarnings.pdf", saveOptions);

            Assert.AreEqual(1, callback.mWarnings.Count);
            Assert.True(callback.mWarnings[0].Description.Contains("R2_XORPEN"));
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
                //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }
        //ExEnd

        [Test]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First)] // Need to check in AW tests
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All)]
        public void HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
        {
            //ExStart
            //ExFor:HeaderFooterBookmarksExportMode
            //ExFor:OutlineOptions
            //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
            //ExSummary:Shows how bookmarks in headers/footers are exported to pdf
            Document doc = new Document(MyDir + "PdfSaveOption.HeaderFooterBookmarksExportMode.docx");

            // You can specify how bookmarks in headers/footers are exported.
            // There is a several options for this:
            // "None" - Bookmarks in headers/footers are not exported.
            // "First" - Only bookmark in first header/footer of the section is exported.
            // "All" - Bookmarks in all headers/footers are exported.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                HeaderFooterBookmarksExportMode = headerFooterBookmarksExportMode,
                OutlineOptions = { DefaultBookmarksOutlineLevel = 1 }
            };
            doc.Save(ArtifactsDir + "PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", saveOptions);
            //ExEnd
        }
    }
}