using Aspose.Words.LowCode;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace PluginsExamples
{
    public class MergerPlugin : PluginsExamplesBase
    {
        [Test]
        public void MergeDocuments()
        {
            //ExStart:MergeDocuments
            //GistId:0dc2066376d6eefd50ebf48307a967ca
            Merger.Merge(ArtifactsDir + "MergerPlugin.MergeDocuments.docx", 
                new[] { MyDir + "Document.docx", MyDir + "Bookmarks.docx" }, 
                new OoxmlSaveOptions() { Password = "Aspose.Words Merger" }, 
                MergeFormatMode.KeepSourceFormatting);
            //ExEnd:MergeDocuments
        }
    }
}
