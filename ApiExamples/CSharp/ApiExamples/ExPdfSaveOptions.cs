// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using NUnit.Framework;
using Document = Aspose.Words.Document;
using IWarningCallback = Aspose.Words.IWarningCallback;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;
using SaveOptions = Aspose.Words.Saving.SaveOptions;
using WarningInfo = Aspose.Words.WarningInfo;
using WarningType = Aspose.Words.WarningType;

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

            doc.Save(MyDir + @"\Artifacts\CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd

            // Bind PDF with Aspose.PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(MyDir + @"\Artifacts\CreateMissingOutlineLevels.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(11, bookmarks.Count);
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

            doc.Save(MyDir + @"\Artifacts\DrawingMl.pdf", pdfSaveOptions);
            //ExEnd
        }

        [Test]
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

            doc.Save(MyDir + @"\Artifacts\UpdateFields_False.pdf", pdfSaveOptions);
            //ExEnd

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + @"\Artifacts\UpdateFields_False.pdf");

            // Get text fragment by search String
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page  of");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are not updated
            Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments[1].Text);
        }

        [Test]
        public void WithUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { UpdateFields = true };

            doc.Save(MyDir + @"\Artifacts\UpdateFields_False.pdf", pdfSaveOptions);

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + @"\Artifacts\UpdateFields_False.pdf");

            // Get text fragment by search String from PDF document
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
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
            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,                
            	JpegQuality = 50 // Use JPEG compression at 50% quality to reduce file size.
            };

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
            //ExEnd

            PdfSaveOptions optionsA1A = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,
                ExportDocumentStructure = true,
                ImageCompression = PdfImageCompression.Jpeg
            };

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageComppression PDF_A_1_A.pdf", optionsA1A);
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

            doc.Save(MyDir + @"\Artifacts\ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
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

            doc.Save(MyDir + @"\Artifacts\PdfTitle.pdf", pdfSaveOptions);
            //ExEnd

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + @"\Artifacts\PdfTitle.pdf");

            Assert.IsTrue(pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
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

            doc.Save(MyDir + @"\Artifacts\SaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        //Note: it works only for spaces. Why?
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

            builder.Document.Save(MyDir + @"\Artifacts\PdfSaveOptions.EscapedUri Out.pdf", options);
            //ExEnd

            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(MyDir + @"\Artifacts\PdfSaveOptions.EscapedUri Out.pdf");

            // get first page
            Page page = pdfDocument.Pages[1];
            // get the first link annotation
            LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

            GoToURIAction action = (GoToURIAction) linkAnnot.Action;
            string uriText = action.URI;

            Assert.AreEqual(result, uriText);
            //ExEnd
        }

        [Test]
        public void HandleBinaryRasterWarnings()
        {
            //ExStart
            //ExFor:MetafileRenderingMode.VectorWithFallback
            //ExFor:IWarningCallback
            //ExFor:PdfSaveOptions.MetafileRenderingOptions
            //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records
            Document doc = new Document(MyDir + "PdfSaveOptions.HandleRasterWarnings.doc");

            MetafileRenderingOptions metafileRenderingOptions =
                new MetafileRenderingOptions
                {
                    EmulateRasterOperations = false,
                    RenderingMode = MetafileRenderingMode.VectorWithFallback
                };

            //If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };
            doc.Save(MyDir + @"\Artifacts\PdfSaveOptions.HandleRasterWarnings.pdf", saveOptions);

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
            doc.Save(MyDir + @"\Artifacts\PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void CreateOutlinesForHeadingsInTables()
        {

        }
    }
}