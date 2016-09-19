// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
            //ExFor:Saving.PdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels
            //ExSummary:Shows how to create missing outline levels saving the document in pdf
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

            //Create "PdfSaveOptions" with some mandatory parameters
            //"HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
            //"CreateMissingOutlineLevels" determining whether or not to create missing heading levels
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(MyDir + @"\Artifacts\CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd

            //Bind pdf with Aspose PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(MyDir + @"\Artifacts\CreateMissingOutlineLevels.pdf");

            //Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(11, bookmarks.Count);
        }

        //Note: Test doesn't containt validation result, because it's difficult
        //For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
        [Test]
        public void DrawingMl()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

            doc.Save(MyDir + @"\Artifacts\DrawingMl.pdf", pdfSaveOptions);
        }

        [Test]
        public void WithoutUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.UpdateFields = false;

            doc.Save(MyDir + @"\Artifacts\UpdateFields_False.pdf", pdfSaveOptions);

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + @"\Artifacts\UpdateFields_False.pdf");

            //Get text fragment by search string
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page  of");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            //Assert that fields are not updated
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

            //Get text fragment by search string
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            //Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
        }

        //For assert this test you need to open "SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf" and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
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
            Document doc = new Document(MyDir + "SaveOptions.PdfImageComppression.rtf");

            PdfSaveOptions options = new PdfSaveOptions();

            options.ImageCompression = PdfImageCompression.Jpeg;
            options.PreserveFormFields = true;

            doc.Save(MyDir + "SaveOptions.PdfImageComppression Out.pdf", options);

            PdfSaveOptions optionsA1b = new PdfSaveOptions();
            optionsA1b.Compliance = PdfCompliance.PdfA1b;
            optionsA1b.ImageCompression = PdfImageCompression.Jpeg;
            optionsA1b.JpegQuality = 100;  // Use JPEG compression at 50% quality to reduce file size.

            doc.Save(MyDir + "SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf", optionsA1b);
            //ExEnd

            PdfSaveOptions optionsA1a = new PdfSaveOptions();
            optionsA1a.Compliance = PdfCompliance.PdfA1a;
            optionsA1a.ExportDocumentStructure = true;
            optionsA1a.ImageCompression = PdfImageCompression.Jpeg;

            doc.Save(MyDir + "SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf", optionsA1a);
            //ExEnd
        }

        [Test]
        public void ColorRendering()
        {
            //ExStart
            //ExFor:PdfSaveOptions.ColorMode
            //ExSummary:Shows how change image color with save options property
            
            //Open document with color image
            Document doc = new Document(MyDir + "Rendering.doc");

            //Set grayscale mode for document
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.ColorMode = ColorMode.Grayscale;

            //Assert that color image in document was grey
            doc.Save(MyDir + @"\Artifacts\ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
            //ExEnd
        }
    }
}
