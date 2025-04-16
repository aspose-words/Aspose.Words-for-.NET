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

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = new PdfEncryptionDetails("Aspose", null);

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", saveOptions);

            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.Password = "Aspose";
            loadOptions.LoadFormat = LoadFormat.Pdf;

            doc = new Document(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", loadOptions);
            //ExEnd:LoadEncryptedPdf
        }

        [Test]
        public void LoadPageRangeOfPdf()
        {
            //ExStart:LoadPageRangeOfPdf  
            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.PageIndex = 0;
            loadOptions.PageCount = 1;

            //ExStart:LoadPDF
            Document doc = new Document(MyDir + "Pdf Document.pdf", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadPageRangeOfPdf.pdf");
            //ExEnd:LoadPDF
            //ExEnd:LoadPageRangeOfPdf
        }
    }
}
