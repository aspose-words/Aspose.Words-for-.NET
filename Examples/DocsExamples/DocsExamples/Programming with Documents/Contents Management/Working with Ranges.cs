using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Management
{
    internal class WorkingWithRanges : DocsExamplesBase
    {
        [Test]
        public void RangesDeleteText()
        {
            //ExStart:RangesDeleteText
            //GistId:9164e9c0658006e51db723b0742c12fc
            Document doc = new Document(MyDir + "Document.docx");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }

        [Test]
        public void RangesGetText()
        {
            //ExStart:RangesGetText
            //GistId:9164e9c0658006e51db723b0742c12fc
            Document doc = new Document(MyDir + "Document.docx");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}