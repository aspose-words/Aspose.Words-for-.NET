using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Managment
{
    internal class WorkingWithRanges : DocsExamplesBase
    {
        [Test]
        public void RangesDeleteText()
        {
            //ExStart:RangesDeleteText
            Document doc = new Document(MyDir + "Document.docx");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
        }

        [Test]
        public void RangesGetText()
        {
            //ExStart:RangesGetText
            Document doc = new Document(MyDir + "Document.docx");
            string text = doc.Range.Text;
            //ExEnd:RangesGetText
        }
    }
}