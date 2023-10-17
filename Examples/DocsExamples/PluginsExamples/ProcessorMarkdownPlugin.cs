using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorMarkdownPlugin : PluginsExamplesBase
    {
        [Test]
        public void AddHorizontalRule()
        {
            //ExStart:AddHorizontalRule
            //GistId:9d4abe412cebe93348409d3632ab3ceb
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert HorizontalRule that will be present in .md file as '-----'.
            builder.InsertHorizontalRule();

            doc.Save(ArtifactsDir + "ProcessorMarkdownPlugin.AddHorizontalRule.md");
            //ExEnd:AddHorizontalRule
        }

        [Test]
        public void EditDocumentMarkdown()
        {
            //ExStart:EditDocumentMarkdown
            //GistId:9d4abe412cebe93348409d3632ab3ceb
            var doc = new Document(MyDir + "Quotes.md");
            var builder = new DocumentBuilder(doc);

            // Prepare created document for further work
            // and clear paragraph formatting not to use the previous styles.
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            // Use optional dot (.) and number of backticks (`).
            // There will be 3 backticks.
            var inlineCode3BackTicks = doc.Styles.Add(StyleType.Character, "InlineCode.3");
            builder.Font.Style = inlineCode3BackTicks;
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorMarkdownPlugin.EditDocumentMarkdown.md");
            //ExEnd:EditDocumentMarkdown
        }
    }
}
