using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithPclSaveOptions : DocsExamplesBase
    {
        [Test]
        public void RasterizeTransformedElements()
        {
            //ExStart:RasterizeTransformedElements
            //GistId:7ee438947078cf070c5bc36a4e45a18c
            Document doc = new Document(MyDir + "Rendering.docx");

            PclSaveOptions saveOptions = new PclSaveOptions
            {
                SaveFormat = SaveFormat.Pcl, RasterizeTransformedElements = false
            };

            doc.Save(ArtifactsDir + "WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
            //ExEnd:RasterizeTransformedElements
        }
    }
}