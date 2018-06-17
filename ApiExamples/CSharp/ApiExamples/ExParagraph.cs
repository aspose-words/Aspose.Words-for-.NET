﻿using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExParagraph : ApiExampleBase
    {
        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:Paragraph.InsertField(string, Node, bool)
            //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
            //ExFor:Paragraph.InsertField(string, string, Node, bool)
            //ExSummary:Shows how to insert field using several methods: "field code", "field code and field value", "field code and field value after a run of text"
            Document doc = new Document();

            // Get the first paragraph of the document
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // Inserting field using field code
            // Note: All methods support inserting field after some node. Just set "true" in the "isAfter" parameter
            para.InsertField(" AUTHOR ", null, false);

            // Using field type
            // Note:
            // 1. For inserting field using field type, you can choose, update field before or after you open the document ("updateField" parameter)
            // 2. For other methods it's works automatically
            para.InsertField(FieldType.FieldAuthor, false, null, true);

            // Using field code and field value
            para.InsertField(" AUTHOR ", "Test Field Value", null, false);

            // Add a run of text
            Run run = new Run(doc) { Text = " Hello World!" };
            para.AppendChild(run);

            // Using field code and field value before a run of text
            // Note: For inserting field before/after a run of text you can use all methods above, just add ref on your text ("refNode" parameter)
            para.InsertField(" AUTHOR ", "Test Field Value", run, false);
            //ExEnd
        }

        [Test]
        public void InsertFieldBeforeTextInParagraph()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterTextInParagraph()
        {
            String date = DateTime.Today.ToString("d");

            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

            Assert.AreEqual(String.Format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date), DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldBeforeTextInParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false, 1);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterTextInParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, true, 1);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldWithoutSeparator()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldListNum, true, null, false, 1);

            Assert.AreEqual("\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldBeforeParagraphWithoutDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            doc.BuiltInDocumentProperties.Author = "";

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldBeforeRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1);

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            // Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1);

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

            Assert.AreEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        [Description("WORDSNET-12396")]
        public void InsertFieldEmptyParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false, 1);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015\f", DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        [Description("WORDSNET-12397")]
        public void InsertFieldEmptyParagraphWithUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, true, null, false, 0);

            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void GetFormatRevision()
        {
            //ExStart
            //ExFor:Paragraph.IsFormatRevision
            //ExSummary:Shows how to get information about whether this object was formatted in Microsoft Word while change tracking was enabled
            Document doc = new Document(MyDir + "Paragraph.IsFormatRevision.docx");

            Paragraph firstParagraph = DocumentHelper.GetParagraph(doc, 0);
            Assert.IsTrue(firstParagraph.IsFormatRevision);
            //ExEnd

            Paragraph secondParagraph = DocumentHelper.GetParagraph(doc, 1);
            Assert.IsFalse(secondParagraph.IsFormatRevision);
        }

        [Test]
        public void GetFrameProperties()
        {
            //ExStart
            //ExFor:Paragraph.FrameFormat
            //ExFor:FrameFormat.IsFrame
            //ExFor:FrameFormat.Width
            //ExFor:FrameFormat.Height
            //ExFor:FrameFormat.HorizontalPosition
            //ExFor:FrameFormat.RelativeHorizontalPosition
            //ExFor:FrameFormat.HorizontalDistanceFromText
            //ExFor:FrameFormat.VerticalPosition
            //ExFor:FrameFormat.RelativeVerticalPosition
            //ExFor:FrameFormat.VerticalDistanceFromText
            //ExSummary:Shows how to get information about formatting properties of paragraph as frame.
            Document doc = new Document(MyDir + "Paragraph.Frame.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.FrameFormat.IsFrame)
                {
                    Console.WriteLine("Width: " + paragraph.FrameFormat.Width);
                    Console.WriteLine("Height: " + paragraph.FrameFormat.Height);
                    Console.WriteLine("HorizontalPosition: " + paragraph.FrameFormat.HorizontalPosition);
                    Console.WriteLine("RelativeHorizontalPosition: " + paragraph.FrameFormat.RelativeHorizontalPosition);
                    Console.WriteLine("HorizontalDistanceFromText: " + paragraph.FrameFormat.HorizontalDistanceFromText);
                    Console.WriteLine("VerticalPosition: " + paragraph.FrameFormat.VerticalPosition);
                    Console.WriteLine("RelativeVerticalPosition: " + paragraph.FrameFormat.RelativeVerticalPosition);
                    Console.WriteLine("VerticalDistanceFromText: " + paragraph.FrameFormat.VerticalDistanceFromText);
                }
            }
            //ExEnd

            if (paragraphs[0].FrameFormat.IsFrame)
            {
                Assert.AreEqual(233.3, paragraphs[0].FrameFormat.Width);
                Assert.AreEqual(138.8, paragraphs[0].FrameFormat.Height);
                Assert.AreEqual(21.05, paragraphs[0].FrameFormat.HorizontalPosition);
                Assert.AreEqual(RelativeHorizontalPosition.Page, paragraphs[0].FrameFormat.RelativeHorizontalPosition);
                Assert.AreEqual(9, paragraphs[0].FrameFormat.HorizontalDistanceFromText);
                Assert.AreEqual(-17.65, paragraphs[0].FrameFormat.VerticalPosition);
                Assert.AreEqual(RelativeVerticalPosition.Paragraph, paragraphs[0].FrameFormat.RelativeVerticalPosition);
                Assert.AreEqual(0, paragraphs[0].FrameFormat.VerticalDistanceFromText);
            }
            else
            {
                Assert.Fail("There are no frames in the document.");    
            }
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field type
        /// </summary>
        private static void InsertFieldUsingFieldType(Document doc, FieldType fieldType, bool updateField, Node refNode, bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldType, updateField, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code
        /// </summary>
        private static void InsertFieldUsingFieldCode(Document doc, String fieldCode, Node refNode, bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code and field String
        /// </summary>
        private static void InsertFieldUsingFieldCodeFieldString(Document doc, String fieldCode, String fieldValue, Node refNode, bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, fieldValue, refNode, isAfter);
        }
    }
}