using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithRtfSaveOptions : DocsExamplesBase
    {
        [Test]
        public void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            //GistId:6f849e51240635a6322ab0460938c922
            Document doc = new Document(MyDir + "Document.docx");

            RtfSaveOptions saveOptions = new RtfSaveOptions { SaveImagesAsWmf = true };

            doc.Save(ArtifactsDir + "WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
            //ExEnd:SavingImagesAsWmf
        }
    }
}