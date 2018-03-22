// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using NUnit.Framework;

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

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

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

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.UpdateFields = false;

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

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.UpdateFields = true;

            doc.Save(MyDir + @"\Artifacts\UpdateFields_False.pdf", pdfSaveOptions);

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + @"\Artifacts\UpdateFields_False.pdf");

            // Get text fragment by search String
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
        }

        //ToDo: Add gold asserts for PDF files
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
            //ExSummary:Demonstrates how to save images to PDF using JPEG encoding to decrease file size.
            Document doc = new Document(MyDir + "SaveOptions.PdfImageCompression.rtf");

            PdfSaveOptions options = new PdfSaveOptions();

            options.ImageCompression = PdfImageCompression.Jpeg;
            options.PreserveFormFields = true;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageCompression Out.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions();
            optionsA1B.Compliance = PdfCompliance.PdfA1b;
            optionsA1B.ImageCompression = PdfImageCompression.Jpeg;
            optionsA1B.JpegQuality = 100; // Use JPEG compression at 50% quality to reduce file size.

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf", optionsA1B);
            //ExEnd

            PdfSaveOptions optionsA1A = new PdfSaveOptions();
            optionsA1A.Compliance = PdfCompliance.PdfA1a;
            optionsA1A.ExportDocumentStructure = true;
            optionsA1A.ImageCompression = PdfImageCompression.Jpeg;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf", optionsA1A);
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
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.ColorMode = ColorMode.Grayscale;

            // Assert that color image in document was grey
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

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.DisplayDocTitle = true;

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

            doc.Save(MyDir + @"\Artifacts\SaveOptions.MemoryOptimization Out.pdf", saveOptions);
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

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
            metafileRenderingOptions.EmulateRasterOperations = false;

            //If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            metafileRenderingOptions.RenderingMode = MetafileRenderingMode.VectorWithFallback;

            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;
            
            doc.Save(MyDir + "PdfSaveOptions.HandleRasterWarnings Out.pdf", saveOptions);

            Assert.AreEqual(1, callback.mWarnings.Count);
            Assert.True(callback.mWarnings[0].Description.Contains("R2_XORPEN"));
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document procssing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    this.mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }//ExEnd
    }
}
