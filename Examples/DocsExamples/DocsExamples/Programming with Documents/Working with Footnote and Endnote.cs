using Aspose.Words;
using Aspose.Words.Notes;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithFootnotes : DocsExamplesBase
    {
        [Test]
        public void SetFootnoteColumns()
        {
            //ExStart:SetFootnoteColumns
            //GistId:3b39c2019380ee905e7d9596494916a4
            Document doc = new Document(MyDir + "Document.docx");

            // Specify the number of columns with which the footnotes area is formatted.
            doc.FootnoteOptions.Columns = 3;
            
            doc.Save(ArtifactsDir + "WorkingWithFootnotes.SetFootnoteColumns.docx");
            //ExEnd:SetFootnoteColumns
        }

        [Test]
        public void SetFootnoteAndEndnotePosition()
        {
            //ExStart:SetFootnoteAndEndnotePosition
            //GistId:3b39c2019380ee905e7d9596494916a4
            Document doc = new Document(MyDir + "Document.docx");

            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;
            
            doc.Save(ArtifactsDir + "WorkingWithFootnotes.SetFootnoteAndEndnotePosition.docx");
            //ExEnd:SetFootnoteAndEndnotePosition
        }

        [Test]
        public void SetEndnoteOptions()
        {
            //ExStart:SetEndnoteOptions
            //GistId:3b39c2019380ee905e7d9596494916a4
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Some text");
            builder.InsertFootnote(FootnoteType.Endnote, "Footnote text.");

            EndnoteOptions option = doc.EndnoteOptions;
            option.RestartRule = FootnoteNumberingRule.RestartPage;
            option.Position = EndnotePosition.EndOfSection;

            doc.Save(ArtifactsDir + "WorkingWithFootnotes.SetEndnoteOptions.docx");
            //ExEnd:SetEndnoteOptions
        }
    }
}