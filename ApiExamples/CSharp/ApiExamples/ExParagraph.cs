using System;
using System.Linq;
using System.Web.UI;
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
            //ExFor:Paragraph.AppendField(FieldType, Boolean)
            //ExFor:Paragraph.AppendField(String)
            //ExFor:Paragraph.AppendField(String, String)
            //ExFor:Paragraph.InsertField(string, Node, bool)
            //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
            //ExFor:Paragraph.InsertField(string, string, Node, bool)
            //ExSummary:Demonstrates various ways of inserting fields.
            // Create a blank document and get its first paragraph
            Document doc = new Document();
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // Choose a field by FieldType, append it to the end of the paragraph and update it
            para.AppendField(FieldType.FieldDate, true);

            // Append a field with a field code created by hand 
            para.AppendField(" TIME  \\@ \"HH:mm:ss\" ");

            // Append a field that will display a placeholder value until it is updated manually in Microsoft Word
            // or programmatically with Document.UpdateFields() or Field.Update()
            para.AppendField(" QUOTE \"Real value\"", "Placeholder value");

            // We can choose a node in the paragraph and insert a field
            // before or after that node instead of appending it to the end of a paragraph
            para = doc.FirstSection.Body.AppendParagraph("");
            Run run = new Run(doc) { Text = " My Run. " };
            para.AppendChild(run);

            // Insert a field into the paragraph and place it before the run we created
            doc.BuiltInDocumentProperties["Author"].Value = "John Doe";
            para.InsertField(FieldType.FieldAuthor, true, run, false);

            // Insert another field designated by field code before the run
            para.InsertField(" QUOTE \"Real value\" ", run, false);

            // Insert another field with a place holder value and place it after the run
            para.InsertField(" QUOTE \"Real value\"", " Placeholder value. ", run, true);

            doc.Save(ArtifactsDir + "Paragraph.InsertField.docx");
            //ExEnd
        }

        [Test]
        public void InsertFieldBeforeTextInParagraph()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r",
                DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterTextInParagraph()
        {
            String date = DateTime.Today.ToString("d");

            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

            Assert.AreEqual(String.Format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date),
                DocumentHelper.GetParagraphText(doc, 1));
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

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r",
                DocumentHelper.GetParagraphText(doc, 1));
        }

        [Test]
        public void InsertFieldAfterRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            // Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1);

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

            Assert.AreEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r",
                DocumentHelper.GetParagraphText(doc, 1));
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
            //ExFor:FrameFormat
            //ExFor:FrameFormat.IsFrame
            //ExFor:FrameFormat.Width
            //ExFor:FrameFormat.Height
            //ExFor:FrameFormat.HeightRule
            //ExFor:FrameFormat.HorizontalAlignment
            //ExFor:FrameFormat.VerticalAlignment
            //ExFor:FrameFormat.HorizontalPosition
            //ExFor:FrameFormat.RelativeHorizontalPosition
            //ExFor:FrameFormat.HorizontalDistanceFromText
            //ExFor:FrameFormat.VerticalPosition
            //ExFor:FrameFormat.RelativeVerticalPosition
            //ExFor:FrameFormat.VerticalDistanceFromText
            //ExSummary:Shows how to get information about formatting properties of paragraph as frame.
            Document doc = new Document(MyDir + "Paragraph.Frame.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            foreach (Paragraph paragraph in paragraphs.OfType<Paragraph>())
            {
                if (paragraph.FrameFormat.IsFrame)
                {
                    Console.WriteLine("Width: " + paragraph.FrameFormat.Width);
                    Console.WriteLine("Height: " + paragraph.FrameFormat.Height);
                    Console.WriteLine("HeightRule: " + paragraph.FrameFormat.HeightRule);
                    Console.WriteLine("HorizontalAlignment: " + paragraph.FrameFormat.HorizontalAlignment);
                    Console.WriteLine("VerticalAlignment: " + paragraph.FrameFormat.VerticalAlignment);
                    Console.WriteLine("HorizontalPosition: " + paragraph.FrameFormat.HorizontalPosition);
                    Console.WriteLine("RelativeHorizontalPosition: " +
                                      paragraph.FrameFormat.RelativeHorizontalPosition);
                    Console.WriteLine("HorizontalDistanceFromText: " +
                                      paragraph.FrameFormat.HorizontalDistanceFromText);
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

        [Test]
        public void AsianTypographyProperties()
        {
            //ExStart
            //ExFor:ParagraphFormat.FarEastLineBreakControl
            //ExFor:ParagraphFormat.WordWrap
            //ExFor:ParagraphFormat.HangingPunctuation
            //ExSummary:Shows how to set special properties for Asian typography. 
            Document doc = new Document(MyDir + "Document.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            format.FarEastLineBreakControl = true;
            format.WordWrap = false;
            format.HangingPunctuation = true;

            doc.Save(ArtifactsDir + "Paragraph.AsianTypographyProperties.docx");
            //ExEnd
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field type
        /// </summary>
        private static void InsertFieldUsingFieldType(Document doc, FieldType fieldType, bool updateField, Node refNode,
            bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldType, updateField, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code
        /// </summary>
        private static void InsertFieldUsingFieldCode(Document doc, String fieldCode, Node refNode, bool isAfter,
            int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code and field String
        /// </summary>
        private static void InsertFieldUsingFieldCodeFieldString(Document doc, String fieldCode, String fieldValue,
            Node refNode, bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, fieldValue, refNode, isAfter);
        }

        [Test]
        public void DropCapPosition()
        {
            //ExStart
            //ExFor:DropCapPosition
            //ExSummary:Shows how to set the position of a drop cap.
            // Create a blank document
            Document doc = new Document();

            // Every paragraph has its own drop cap setting
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // By default, it is "none", for no drop caps
            Assert.AreEqual(Aspose.Words.DropCapPosition.None, para.ParagraphFormat.DropCapPosition);

            // Move the first capital to outside the text margin
            para.ParagraphFormat.DropCapPosition = Aspose.Words.DropCapPosition.Margin;

            // This text will be affected
            para.Runs.Add(new Run(doc, "Hello World!"));

            doc.Save(ArtifactsDir + "Paragraph.DropCap.docx");
            //ExEnd
        }

        [Test]
        public void IsRevision()
        {
            //ExStart
            //ExFor:Paragraph.IsDeleteRevision
            //ExFor:Paragraph.IsInsertRevision
            //ExSummary:Shows how to work with revision paragraphs
            // Create a blank document, populate the first paragraph with text and add two more
            Document doc = new Document();
            Body body = doc.FirstSection.Body;
            Paragraph para = body.FirstParagraph;
            para.AppendChild(new Run(doc, "Paragraph 1. "));
            body.AppendParagraph("Paragraph 2. ");
            body.AppendParagraph("Paragraph 3. ");

            // We have three paragraphs, none of which registered as any type of revision
            // If we add/remove any content in the document while tracking revisions,
            // they will be displayed as such in the document and can be accepted/rejected
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            // This paragraph is a revision and will have the according "IsInsertRevision" flag set
            para = body.AppendParagraph("Paragraph 4. ");
            Assert.True(para.IsInsertRevision);
            
            // Get the document's paragraph collection and remove a paragraph
            ParagraphCollection paragraphs = body.Paragraphs;
            Assert.AreEqual(4, paragraphs.Count);
            para = paragraphs[2];
            para.Remove();

            // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
            // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions
            Assert.AreEqual(4, paragraphs.Count);
            Assert.True(para.IsDeleteRevision);

            // The delete revision paragraph is removed once we accept changes
            doc.AcceptAllRevisions();
            Assert.AreEqual(3, paragraphs.Count);
            Assert.IsEmpty(para);
            //ExEnd
        }

        [Test]
        public void BreakIsStyleSeparator()
        {
            //ExStart
            //ExFor:Paragraph.BreakIsStyleSeparator
            //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
            // Create a blank document and insert a table of contents field
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTableOfContents("\\o \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Insert a paragraph with a style that will be picked up as an entry in the TOC
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Both these strings are on the same line and same paragraph and will therefore show up on the same TOC entry
            builder.Write("Heading 1. ");
            builder.Write("Will appear in the TOC. ");

            // Any text on a new line that does not have a heading style will not register as a TOC entry
            // If we insert a style separator, we can write more text on the same line
            // and use a different style without it showing up in the TOC
            // If we use a heading type style afterwards, we can draw two TOC entries from one line of document text
            builder.InsertStyleSeparator();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
            builder.Write("Won't appear in the TOC. ");

            // This flag is set to true for such paragraphs
            Assert.True(doc.FirstSection.Body.Paragraphs[0].BreakIsStyleSeparator);

            // Update the TOC and save the document
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Paragraph.BreakIsStyleSeparator.docx");
            //ExEnd
        }
    }
}