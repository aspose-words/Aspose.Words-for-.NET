using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    public class WorkingWithPdfLoadOptions : DocsExamplesBase
    {
        [Test]
        public void LoadEncryptedPdf()
        {
            //ExStart:LoadEncryptedPdf  
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                EncryptionDetails = new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC4_40)
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", saveOptions);

            PdfLoadOptions loadOptions = new PdfLoadOptions { Password = "Aspose", LoadFormat = LoadFormat.Pdf };

            doc = new Document(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", loadOptions);
            //ExEnd:LoadEncryptedPdf
        }

        [Test]
        public void LoadPageRangeOfPdf()
        {
            //ExStart:LoadPageRangeOfPdf  
            PdfLoadOptions loadOptions = new PdfLoadOptions { PageIndex = 0, PageCount = 1 };

            //ExStart:LoadPDF
            Document doc = new Document(MyDir + "Pdf Document.pdf", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadPageRangeOfPdf.pdf");
            //ExEnd:LoadPDF
            //ExEnd:LoadPageRangeOfPdf
        }
    }
}
