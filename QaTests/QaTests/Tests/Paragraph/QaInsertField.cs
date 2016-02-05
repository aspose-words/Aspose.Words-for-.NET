using System;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace QaTests.Tests
{
    /// <summary>
    /// Tests that verify inserting field into the paragraph
    /// </summary>
    [TestFixture]
    class QaInsertField : QaTestsBase
    {
        [Test]
        public void InsertField_BeforeTextInParagraph()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " AUTHOR ", null, false);
            
            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_AfterTextInParagraph()
        {
            string date = DateTime.Today.ToString("d");

            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " DATE ", null, true);

            Assert.AreEqual(String.Format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date), DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_BeforeTextInParagraph_WithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_AfterTextInParagraph_WithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, true);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldWithoutSeparator()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldListNum, true, null, false);

            Assert.AreEqual("\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_BeforeParagraph_WithoutDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            doc.BuiltInDocumentProperties.Author = "";

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_AfterParagraph_WithoutChangingDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_BeforeRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!");

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertField_AfterRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!");

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true);

            Assert.AreEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Test for WORDSNET-12396
        /// </summary>
        [Test]
        public void InsertField_EmptyParagraph_WithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Test for WORDSNET-12397
        /// </summary>
        [Test]
        public void InsertField_EmptyParagraph_WithUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, true, null, false);
            
            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field type
        /// </summary>
        private static void InsertFieldUsingFieldType(Document doc, FieldType fieldType, bool updateField, Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldType, updateField, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code
        /// </summary>
        private static void InsertFieldUsingFieldCode(Document doc, string fieldCode, Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldCode, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code and field string
        /// </summary>
        private static void InsertFieldUsingFieldCodeFieldString(Document doc, string fieldCode, string fieldValue, Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldCode, fieldValue, refNode, isAfter);
        }
    }
}
