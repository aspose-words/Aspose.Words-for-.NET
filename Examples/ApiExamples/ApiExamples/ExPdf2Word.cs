using NUnit.Framework;
using Aspose.Words;

#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf.Text;
#endif

namespace ApiExamples
{
#if NET462 || NETCOREAPP2_1 || JAVA
    [TestFixture]
    public class ExPdf2Word : ApiExampleBase
    {
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
            Aspose.Words.Document loadedPdfAw = new Aspose.Words.Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            Assert.AreEqual("Hello world!", loadedPdfAw.GetText().Trim());

            // 2 -  Load as an Aspose.Pdf document:
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PDF2Word.LoadPdf.pdf");

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            Assert.AreEqual("Hello world!", textFragmentAbsorber.Text.Trim());
            //ExEnd
        }
    }
#endif
}
