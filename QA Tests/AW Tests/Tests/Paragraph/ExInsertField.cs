using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace QA_Tests.Tests
{
    [TestFixture]
    internal class ExInsertField : QaTestsBase
    {
        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:Paragraph.InsertField
            //ExSummary:Shows how to insert field using several methods: "field code", "field code and field value", "field code and field value after a run of text"
            Document doc = new Document();

            //Get the first paragraph of the document
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            //Inseting field using field code
            //Note: All methods support inserting field after some node. Just set "true" in the "isAfter" parameter
            para.InsertField(" AUTHOR ", null, false);

            //Using field type
            //Note:
            //1. For inserting field using field type, you can choose, update field before or after you open the document ("updateField" parameter)
            //2. For other methods it's works automatically
            para.InsertField(FieldType.FieldAuthor, false, null, true);

            //Using field code and field value
            para.InsertField(" AUTHOR ", "Test Field Value", null, false);

            //Add a run of text
            Run run = new Run(doc) { Text = " Hello World!" };
            para.AppendChild(run);

            //Using field code and field value before a run of text
            //Note: For inserting field before/after a run of text you can use all methods above, just add ref on your text ("refNode" parameter)
            para.InsertField(" AUTHOR ", "Test Field Value", run, false);
            //ExEnd
        }
    }
}
