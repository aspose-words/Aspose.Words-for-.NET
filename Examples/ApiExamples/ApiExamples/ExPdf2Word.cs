using NUnit.Framework;
using Aspose.Words;

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
    }
}
