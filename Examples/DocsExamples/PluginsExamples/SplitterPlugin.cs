using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class SplitterPlugin : PluginsExamplesBase
    {
        [Test]
        public void SplitDocument()
        {
            //ExStart:SplitDocument
            //GistId:22ae7036961fe3bde53ad6802c15edeb
            var doc = new Document(MyDir + "Big document.docx");

            for (var page = 0; page < doc.PageCount; page++)
            {
                var extractedPage = doc.ExtractPages(page, 1);
                extractedPage.Save(ArtifactsDir + $"SplitterPlugin.SplitDocument_{page + 1}.docx");
            }
            //ExEnd:SplitDocument
        }
    }
}
