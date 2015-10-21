using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace QA_Tests.Tests
{
    [TestFixture]
    internal class ExInsertField : QaTestsBase
    {
        [Test]
        public void InsertField_FieldCode()
        {
            //ExStart
            //ExFor:Paragraph
            //ExFor:Paragraph.InsertField
            //ExFor:Field
            //ExSummary:Shows how to insert field using field code into the first paragraph
            Document doc = new Document();

            //Insert field using field code into the first paragraph
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.InsertField(" AUTHOR ", null, false);
            //ExEnd
        }

        [Test]
        public void InsertField_FieldType()
        {
            //ExStart
            //ExFor:Paragraph
            //ExFor:Paragraph.InsertField
            //ExFor:Field
            //ExSummary:Shows how to insert field using field code into the first paragraph
            Document doc = new Document();

            //Insert field using field type into the first paragraph
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.InsertField(FieldType.FieldAuthor, false, null, false);
            //ExEnd
        }

        [Test]
        public void InsertField_FieldCodeAndFieldValue()
        {
            //ExStart
            //ExFor:Paragraph
            //ExFor:Paragraph.InsertField
            //ExFor:Field
            //ExSummary:Shows how to insert field using field code and field value into the first paragraph
            Document doc = new Document();

            //Insert field using field code and field value into the first paragraph
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.InsertField(" AUTHOR ", "Test Field Value", null, false);
            //ExEnd
        }

        [Test]
        public void InsertField_RunText()
        {
            //ExStart
            //ExFor:Paragraph
            //ExFor:Paragraph.InsertField
            //ExFor:Field
            //ExSummary:Shows how to insert field befor/after a run of text
            Document doc = new Document();

            //Get the first paragraph of the document
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            //Add a run of text
            Run run = new Run(doc) { Text = " Hello World!" };
            para.AppendChild(run);

            //Insert field befor a run of text
            //For inserting field after a run of text, you must set "true" in the "isAfter" parameter
            para.InsertField(" AUTHOR ", "Test Field Value", run, false);
            //ExEnd
        }
    }
}
