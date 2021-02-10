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
            Document doc = new Document(MyDir + "Document.docx");

            RtfSaveOptions saveOptions = new RtfSaveOptions { SaveImagesAsWmf = true };

            doc.Save(ArtifactsDir + "WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
            //ExEnd:SavingImagesAsWmf
        }
    }
}