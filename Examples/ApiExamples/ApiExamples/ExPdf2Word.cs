﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using NUnit.Framework;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Pdf.Text;

namespace ApiExamples
{
    [TestFixture]
    public class ExPdf2Word : ApiExampleBase
    {
        [Test]
        public void LoadPdf()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExSummary:Shows how to load a PDF.
            Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            // Below are two ways of loading PDF documents using Aspose products.
            // 1 -  Load as an Aspose.Words document:
            Document asposeWordsDoc = new Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            Assert.That(asposeWordsDoc.GetText().Trim(), Is.EqualTo("Hello world!"));

            // 2 -  Load as an Aspose.Pdf document:
            Aspose.Pdf.Document asposePdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            asposePdfDoc.Pages.Accept(textFragmentAbsorber);

            Assert.That(textFragmentAbsorber.Text.Trim(), Is.EqualTo("Hello world!"));
            //ExEnd
        }

        [Test]
        public static void ConvertPdfToDocx()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExFor:Document.Save(String)
            //ExSummary:Shows how to convert a PDF to a .docx.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "PDF2Word.ConvertPdfToDocx.pdf");

            // Load the PDF document that we just saved, and convert it to .docx.
            Document pdfDoc = new Document(ArtifactsDir + "PDF2Word.ConvertPdfToDocx.pdf");

            pdfDoc.Save(ArtifactsDir + "PDF2Word.ConvertPdfToDocx.docx");
            //ExEnd
        }

        [Test]
        public static void ConvertPdfToDocxCustom()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExSummary:Shows how to convert a PDF to a .docx and customize the saving process with a SaveOptions object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            doc.Save(ArtifactsDir + "PDF2Word.ConvertPdfToDocxCustom.pdf");

            // Load the PDF document that we just saved, and convert it to .docx.
            Document pdfDoc = new Document(ArtifactsDir + "PDF2Word.ConvertPdfToDocxCustom.pdf");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

            // Set the "Password" property to encrypt the saved document with a password.
            saveOptions.Password = "MyPassword";

            pdfDoc.Save(ArtifactsDir + "PDF2Word.ConvertPdfToDocxCustom.docx", saveOptions);
            //ExEnd
        }

        [Test]
        public static void LoadEncryptedPdf()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world! This is an encrypted PDF document.");

            // Configure a SaveOptions object to encrypt this PDF document while saving it to the local file system.
            PdfEncryptionDetails encryptionDetails =
                new PdfEncryptionDetails("MyPassword", string.Empty);

            Assert.That(encryptionDetails.Permissions, Is.EqualTo(PdfPermissions.DisallowAll));

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = encryptionDetails;

            doc.Save(ArtifactsDir + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", saveOptions);

            // To load a password encrypted document, we need to pass a LoadOptions object
            // with the correct password stored in its "Password" property.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyPassword";

            Document pdfDoc = new Document(ArtifactsDir + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", loadOptions);

            Assert.That(pdfDoc.GetText().Trim(), Is.EqualTo("Hello world! This is an encrypted PDF document."));
        }
    }
}
