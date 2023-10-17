using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace PluginsExamples
{
    public class ProcessorTextPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateDocumentText()
        {
            //ExStart:CreateDocumentText
            //GistId:1d6193fa1b96defb817c8d4aa63e00f1
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");
            builder.Write("Paragraph 3.");

            // Set the "ParagraphBreak" to a custom value that 
            // we wish to put at the end of every paragraph.
            var txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ParagraphBreak = " End of paragraph.\n\n\t";

            doc.Save(ArtifactsDir + "ProcessorTextPlugin.CreateDocumentText.txt");
            //ExEnd:CreateDocumentText
        }

        [Test]
        public void EditDocumentText()
        {
            //ExStart:EditDocumentText
            //GistId:1d6193fa1b96defb817c8d4aa63e00f1
            var doc = new Document(MyDir + "English text.txt");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorTextPlugin.EditDocumentText.txt");
            //ExEnd:EditDocumentText
        }
    }
}
