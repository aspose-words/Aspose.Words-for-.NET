using System.IO;
using NUnit.Framework;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf.Text;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExPdf2Word : ApiExampleBase
    {
#if NET462 || NETCOREAPP2_1 || JAVA
        [Test]
        public void LoadPdf()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExSummary:Shows how to load a PDF.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            // Below are two ways of loading PDF documents using Aspose products.
            // 1 -  Load as an Aspose.Words document:
            Aspose.Words.Document asposeWordsDoc = new Aspose.Words.Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            Assert.AreEqual("Hello world!", asposeWordsDoc.GetText().Trim());

            // 2 -  Load as an Aspose.Pdf document:
            Aspose.Pdf.Document asposePdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            asposePdfDoc.Pages.Accept(textFragmentAbsorber);

            Assert.AreEqual("Hello world!", textFragmentAbsorber.Text.Trim());
            //ExEnd
        }
#endif

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
        public static void LoadPdfUsingPlugin()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "PDF2Word.LoadPdfUsingPlugin.pdf");

            // Use the Pdf2Word plugin to open load a PDF document as an Aspose.Words document.
            Document pdfDoc = new Document();

            Aspose.Words.Pdf2Word.PdfDocumentReaderPlugin pdf2Word = new Aspose.Words.Pdf2Word.PdfDocumentReaderPlugin();
            using (FileStream stream =
                new FileStream(ArtifactsDir + "PDF2Word.LoadPdfUsingPlugin.pdf", FileMode.Open))
            {
                pdf2Word.Read(stream, new LoadOptions(), pdfDoc);
            }

            builder = new DocumentBuilder(pdfDoc);

            builder.MoveToDocumentEnd();
            builder.Writeln(" We are editing a PDF document that was loaded into Aspose.Words!");

            Assert.AreEqual("Hello world! We are editing a PDF document that was loaded into Aspose.Words!", 
                pdfDoc.GetText().Trim());
        }

        [Test]
        public static void LoadEncryptedPdfUsingPlugin()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world! This is an encrypted PDF document.");

            // Configure a SaveOptions object to encrypt this PDF document while saving it to the local file system.
            PdfEncryptionDetails encryptionDetails =
                new PdfEncryptionDetails("MyPassword", string.Empty, PdfEncryptionAlgorithm.RC4_128);

            Assert.AreEqual(PdfPermissions.DisallowAll, encryptionDetails.Permissions);

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = encryptionDetails;

            doc.Save(ArtifactsDir + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", saveOptions);

            Document pdfDoc = new Document();

            // To load a password encrypted document, we need to pass a LoadOptions object
            // with the correct password stored in its "Password" property.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyPassword";

            Aspose.Words.Pdf2Word.PdfDocumentReaderPlugin pdf2Word = new Aspose.Words.Pdf2Word.PdfDocumentReaderPlugin();
            using (FileStream stream =
                new FileStream(ArtifactsDir + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", FileMode.Open))
            {
                // Pass the LoadOptions object into the Pdf2Word plugin's "Read" method
                // the same way we would pass it into a document's "Load" method.
                pdf2Word.Read(stream, new LoadOptions("MyPassword"), pdfDoc);
            }

            Assert.AreEqual("Hello world! This is an encrypted PDF document.",
                pdfDoc.GetText().Trim());
        }
    }
}
