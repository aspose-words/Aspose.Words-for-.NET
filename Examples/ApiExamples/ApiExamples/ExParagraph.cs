﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using NUnit.Framework;
using Font = Aspose.Words.Font;

namespace ApiExamples
{
    [TestFixture]
    internal class ExParagraph : ApiExampleBase
    {
        [Test]
        public void DocumentBuilderInsertParagraph()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertParagraph
            //ExFor:ParagraphFormat.FirstLineIndent
            //ExFor:ParagraphFormat.Alignment
            //ExFor:ParagraphFormat.KeepTogether
            //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndAlpha
            //ExFor:ParagraphFormat.AddSpaceBetweenFarEastAndDigit
            //ExFor:Paragraph.IsEndOfDocument
            //ExSummary:Shows how to insert a paragraph into the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.AddSpaceBetweenFarEastAndAlpha = true;
            paragraphFormat.AddSpaceBetweenFarEastAndDigit = true;
            paragraphFormat.KeepTogether = true;

            // The "Writeln" method ends the paragraph after appending text
            // and then starts a new line, adding a new paragraph.
            builder.Writeln("Hello world!");

            Assert.That(builder.CurrentParagraph.IsEndOfDocument, Is.True);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            Assert.That(paragraph.ParagraphFormat.FirstLineIndent, Is.EqualTo(8));
            Assert.That(paragraph.ParagraphFormat.Alignment, Is.EqualTo(ParagraphAlignment.Justify));
            Assert.That(paragraph.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha, Is.True);
            Assert.That(paragraph.ParagraphFormat.AddSpaceBetweenFarEastAndDigit, Is.True);
            Assert.That(paragraph.ParagraphFormat.KeepTogether, Is.True);
            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Hello world!"));

            Font runFont = paragraph.Runs[0].Font;

            Assert.That(runFont.Size, Is.EqualTo(16.0d));
            Assert.That(runFont.Bold, Is.True);
            Assert.That(runFont.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            Assert.That(runFont.Name, Is.EqualTo("Arial"));
            Assert.That(runFont.Underline, Is.EqualTo(Underline.Dash));
        }

        [Test]
        public void AppendField()
        {
            //ExStart
            //ExFor:Paragraph.AppendField(FieldType, Boolean)
            //ExFor:Paragraph.AppendField(String)
            //ExFor:Paragraph.AppendField(String, String)
            //ExSummary:Shows various ways of appending fields to a paragraph.
            Document doc = new Document();
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // Below are three ways of appending a field to the end of a paragraph.
            // 1 -  Append a DATE field using a field type, and then update it:
            paragraph.AppendField(FieldType.FieldDate, true);

            // 2 -  Append a TIME field using a field code: 
            paragraph.AppendField(" TIME  \\@ \"HH:mm:ss\" ");

            // 3 -  Append a QUOTE field using a field code, and get it to display a placeholder value:
            paragraph.AppendField(" QUOTE \"Real value\"", "Placeholder value");

            Assert.That(doc.Range.Fields[2].Result, Is.EqualTo("Placeholder value"));

            // This field will display its placeholder value until we update it.
            doc.UpdateFields();

            Assert.That(doc.Range.Fields[2].Result, Is.EqualTo("Real value"));

            doc.Save(ArtifactsDir + "Paragraph.AppendField.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Paragraph.AppendField.docx");

            TestUtil.VerifyField(FieldType.FieldDate, " DATE ", DateTime.Now, doc.Range.Fields[0], new TimeSpan(0, 0, 0, 0));
            TestUtil.VerifyField(FieldType.FieldTime, " TIME  \\@ \"HH:mm:ss\" ", DateTime.Now, doc.Range.Fields[1], new TimeSpan(0, 0, 0, 5));
            TestUtil.VerifyField(FieldType.FieldQuote, " QUOTE \"Real value\"", "Real value", doc.Range.Fields[2]);
        }

        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:Paragraph.InsertField(string, Node, bool)
            //ExFor:Paragraph.InsertField(FieldType, bool, Node, bool)
            //ExFor:Paragraph.InsertField(string, string, Node, bool)
            //ExSummary:Shows various ways of adding fields to a paragraph.
            Document doc = new Document();
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // Below are three ways of inserting a field into a paragraph.
            // 1 -  Insert an AUTHOR field into a paragraph after one of the paragraph's child nodes:
            Run run = new Run(doc) { Text = "This run was written by " };
            para.AppendChild(run);

            doc.BuiltInDocumentProperties["Author"].Value = "John Doe";
            para.InsertField(FieldType.FieldAuthor, true, run, true);

            // 2 -  Insert a QUOTE field after one of the paragraph's child nodes:
            run = new Run(doc) { Text = "." };
            para.AppendChild(run);

            Field field = para.InsertField(" QUOTE \" Real value\" ", run, true);

            // 3 -  Insert a QUOTE field before one of the paragraph's child nodes,
            // and get it to display a placeholder value:
            para.InsertField(" QUOTE \" Real value.\"", " Placeholder value.", field.Start, false);

            Assert.That(doc.Range.Fields[1].Result, Is.EqualTo(" Placeholder value."));

            // This field will display its placeholder value until we update it.
            doc.UpdateFields();

            Assert.That(doc.Range.Fields[1].Result, Is.EqualTo(" Real value."));

            doc.Save(ArtifactsDir + "Paragraph.InsertField.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Paragraph.InsertField.docx");

            TestUtil.VerifyField(FieldType.FieldAuthor, " AUTHOR ", "John Doe", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldQuote, " QUOTE \" Real value.\"", " Real value.", doc.Range.Fields[1]);
            TestUtil.VerifyField(FieldType.FieldQuote, " QUOTE \" Real value\" ", " Real value", doc.Range.Fields[2]);
        }

        [Test]
        public void InsertFieldBeforeTextInParagraph()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r"));
        }

        [Test]
        public void InsertFieldAfterTextInParagraph()
        {
            string date = DateTime.Today.ToString("d");

            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo(string.Format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date)));
        }

        [Test]
        public void InsertFieldBeforeTextInParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("\u0013 AUTHOR \u0014\u0015Hello World!\r"));
        }

        [Test]
        public void InsertFieldAfterTextInParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, true, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("Hello World!\u0013 AUTHOR \u0014\u0015\r"));
        }

        [Test]
        public void InsertFieldWithoutSeparator()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldListNum, true, null, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("\u0013 LISTNUM \u0015Hello World!\r"));
        }

        [Test]
        public void InsertFieldBeforeParagraphWithoutDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            doc.BuiltInDocumentProperties.Author = "";

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("\u0013 AUTHOR \u0014\u0015Hello World!\r"));
        }

        [Test]
        public void InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("Hello World!\u0013 AUTHOR \u0014\u0015\r"));
        }

        [Test]
        public void InsertFieldBeforeRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1);

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r"));
        }

        [Test]
        public void InsertFieldAfterRunText()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            // Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1);

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r"));
        }

        [Test]
        [Description("WORDSNET-12396")]
        public void InsertFieldEmptyParagraphWithoutUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false, 1);

            Assert.That(DocumentHelper.GetParagraphText(doc, 1), Is.EqualTo("\u0013 AUTHOR \u0014\u0015\f"));
        }

        [Test]
        [Description("WORDSNET-12397")]
        public void InsertFieldEmptyParagraphWithUpdateField()
        {
            Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, true, null, false, 0);

            Assert.That(DocumentHelper.GetParagraphText(doc, 0), Is.EqualTo("\u0013 AUTHOR \u0014Test Author\u0015\r"));
        }

        [Test]
        public void CompositeNodeChildren()
        {
            //ExStart
            //ExFor:CompositeNode.Count
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExFor:CompositeNode.InsertAfter``1(``0, Node)
            //ExFor:CompositeNode.InsertBefore``1(``0, Node)
            //ExFor:CompositeNode.PrependChild``1(``0)
            //ExFor:Paragraph.GetText
            //ExFor:Run
            //ExSummary:Shows how to add, update and delete child nodes in a CompositeNode's collection of children.
            Document doc = new Document();

            // An empty document, by default, has one paragraph.
            Assert.That(doc.FirstSection.Body.Paragraphs.Count, Is.EqualTo(1));

            // Composite nodes such as our paragraph can contain other composite and inline nodes as children.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            Run paragraphText = new Run(doc, "Initial text. ");
            paragraph.AppendChild(paragraphText);

            // Create three more run nodes.
            Run run1 = new Run(doc, "Run 1. ");
            Run run2 = new Run(doc, "Run 2. ");
            Run run3 = new Run(doc, "Run 3. ");

            // The document body will not display these runs until we insert them into a composite node
            // that itself is a part of the document's node tree, as we did with the first run.
            // We can determine where the text contents of nodes that we insert
            // appears in the document by specifying an insertion location relative to another node in the paragraph.
            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Initial text."));

            // Insert the second run into the paragraph in front of the initial run.
            paragraph.InsertBefore(run2, paragraphText);

            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Run 2. Initial text."));

            // Insert the third run after the initial run.
            paragraph.InsertAfter(run3, paragraphText);

            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Run 2. Initial text. Run 3."));

            // Insert the first run to the start of the paragraph's child nodes collection.
            paragraph.PrependChild(run1);

            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Run 1. Run 2. Initial text. Run 3."));
            Assert.That(paragraph.GetChildNodes(NodeType.Any, true).Count, Is.EqualTo(4));

            // We can modify the contents of the run by editing and deleting existing child nodes.
            ((Run)paragraph.GetChildNodes(NodeType.Run, true)[1]).Text = "Updated run 2. ";
            paragraph.GetChildNodes(NodeType.Run, true).Remove(paragraphText);

            Assert.That(paragraph.GetText().Trim(), Is.EqualTo("Run 1. Updated run 2. Run 3."));
            Assert.That(paragraph.GetChildNodes(NodeType.Any, true).Count, Is.EqualTo(3));
            //ExEnd
        }

        [Test]
        public void MoveRevisions()
        {
            //ExStart
            //ExFor:Paragraph.IsMoveFromRevision
            //ExFor:Paragraph.IsMoveToRevision
            //ExFor:ParagraphCollection
            //ExFor:ParagraphCollection.Item(Int32)
            //ExFor:Story.Paragraphs
            //ExSummary:Shows how to check whether a paragraph is a move revision.
            Document doc = new Document(MyDir + "Revisions.docx");

            // This document contains "Move" revisions, which appear when we highlight text with the cursor,
            // and then drag it to move it to another location
            // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
            Assert.That(doc.Revisions.Count(r => r.RevisionType == RevisionType.Moving), Is.EqualTo(6));

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Move revisions consist of pairs of "Move from", and "Move to" revisions. 
            // These revisions are potential changes to the document that we can either accept or reject.
            // Before we accept/reject a move revision, the document
            // must keep track of both the departure and arrival destinations of the text.
            // The second and the fourth paragraph define one such revision, and thus both have the same contents.
            Assert.That(paragraphs[3].GetText(), Is.EqualTo(paragraphs[1].GetText()));

            // The "Move from" revision is the paragraph where we dragged the text from.
            // If we accept the revision, this paragraph will disappear,
            // and the other will remain and no longer be a revision.
            Assert.That(paragraphs[1].IsMoveFromRevision, Is.True);

            // The "Move to" revision is the paragraph where we dragged the text to.
            // If we reject the revision, this paragraph instead will disappear, and the other will remain.
            Assert.That(paragraphs[3].IsMoveToRevision, Is.True);
            //ExEnd
        }

        [Test]
        public void RangeRevisions() 
        {
            //ExStart
            //ExFor:Range.Revisions
            //ExSummary:Shows how to work with revisions in range.
            Document doc = new Document(MyDir + "Revisions.docx");

            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            foreach (Revision revision in paragraph.Range.Revisions)
            {
                if (revision.RevisionType == RevisionType.Deletion)
                    revision.Accept();
            }

            // Reject the first section revisions.
            doc.FirstSection.Range.Revisions.RejectAll();
            //ExEnd
        }

        [Test]
        public void GetFormatRevision()
        {
            //ExStart
            //ExFor:Paragraph.IsFormatRevision
            //ExSummary:Shows how to check whether a paragraph is a format revision.
            Document doc = new Document(MyDir + "Format revision.docx");

            // This paragraph is a "Format" revision, which occurs when we change the formatting of existing text
            // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
            Assert.That(doc.FirstSection.Body.FirstParagraph.IsFormatRevision, Is.True);
            //ExEnd
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
            //ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
            Document doc = new Document(MyDir + "Paragraph frame.docx");

            Paragraph paragraphFrame = doc.FirstSection.Body.Paragraphs.OfType<Paragraph>().First(p => p.FrameFormat.IsFrame);

            Assert.That(paragraphFrame.FrameFormat.Width, Is.EqualTo(233.3d));
            Assert.That(paragraphFrame.FrameFormat.Height, Is.EqualTo(138.8d));
            Assert.That(paragraphFrame.FrameFormat.HeightRule, Is.EqualTo(HeightRule.AtLeast));
            Assert.That(paragraphFrame.FrameFormat.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Default));
            Assert.That(paragraphFrame.FrameFormat.VerticalAlignment, Is.EqualTo(VerticalAlignment.Default));
            Assert.That(paragraphFrame.FrameFormat.HorizontalPosition, Is.EqualTo(34.05d));
            Assert.That(paragraphFrame.FrameFormat.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Page));
            Assert.That(paragraphFrame.FrameFormat.HorizontalDistanceFromText, Is.EqualTo(9.0d));
            Assert.That(paragraphFrame.FrameFormat.VerticalPosition, Is.EqualTo(20.5d));
            Assert.That(paragraphFrame.FrameFormat.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));
            Assert.That(paragraphFrame.FrameFormat.VerticalDistanceFromText, Is.EqualTo(0.0d));
            //ExEnd
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field type.
        /// </summary>
        private static void InsertFieldUsingFieldType(Document doc, FieldType fieldType, bool updateField, Node refNode,
            bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldType, updateField, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code.
        /// </summary>
        private static void InsertFieldUsingFieldCode(Document doc, string fieldCode, Node refNode, bool isAfter,
            int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code and field String.
        /// </summary>
        private static void InsertFieldUsingFieldCodeFieldString(Document doc, string fieldCode, string fieldValue,
            Node refNode, bool isAfter, int paraIndex)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, paraIndex);
            para.InsertField(fieldCode, fieldValue, refNode, isAfter);
        }

        [Test]
        public void IsRevision()
        {
            //ExStart
            //ExFor:Paragraph.IsDeleteRevision
            //ExFor:Paragraph.IsInsertRevision
            //ExSummary:Shows how to work with revision paragraphs.
            Document doc = new Document();
            Body body = doc.FirstSection.Body;
            Paragraph para = body.FirstParagraph;

            para.AppendChild(new Run(doc, "Paragraph 1. "));
            body.AppendParagraph("Paragraph 2. ");
            body.AppendParagraph("Paragraph 3. ");

            // The above paragraphs are not revisions.
            // Paragraphs that we add after starting revision tracking will register as "Insert" revisions.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            para = body.AppendParagraph("Paragraph 4. ");

            Assert.That(para.IsInsertRevision, Is.True);

            // Paragraphs that we remove after starting revision tracking will register as "Delete" revisions.
            ParagraphCollection paragraphs = body.Paragraphs;

            Assert.That(paragraphs.Count, Is.EqualTo(4));

            para = paragraphs[2];
            para.Remove();

            // Such paragraphs will remain until we either accept or reject the delete revision.
            // Accepting the revision will remove the paragraph for good,
            // and rejecting the revision will leave it in the document as if we never deleted it.
            Assert.That(paragraphs.Count, Is.EqualTo(4));
            Assert.That(para.IsDeleteRevision, Is.True);

            // Accept the revision, and then verify that the paragraph is gone.
            doc.AcceptAllRevisions();

            Assert.That(paragraphs.Count, Is.EqualTo(3));
            Assert.That(para.Count, Is.EqualTo(0));
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Paragraph 1. \r" +
                "Paragraph 2. \r" +
                "Paragraph 4."));
            //ExEnd
        }

        [Test]
        public void BreakIsStyleSeparator()
        {
            //ExStart
            //ExFor:Paragraph.BreakIsStyleSeparator
            //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertTableOfContents("\\o \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Insert a paragraph with a style that the TOC will pick up as an entry.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Both these strings are in the same paragraph and will therefore show up on the same TOC entry.
            builder.Write("Heading 1. ");
            builder.Write("Will appear in the TOC. ");

            // If we insert a style separator, we can write more text in the same paragraph
            // and use a different style without showing up in the TOC.
            // If we use a heading type style after the separator, we can draw multiple TOC entries from one document text line.
            builder.InsertStyleSeparator();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
            builder.Write("Won't appear in the TOC. ");

            Assert.That(doc.FirstSection.Body.FirstParagraph.BreakIsStyleSeparator, Is.True);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Paragraph.BreakIsStyleSeparator.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Paragraph.BreakIsStyleSeparator.docx");

            TestUtil.VerifyField(FieldType.FieldTOC, "TOC \\o \\h \\z \\u", 
                "\u0013 HYPERLINK \\l \"_Toc256000000\" \u0014Heading 1. Will appear in the TOC.\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\u0015\r", doc.Range.Fields[0]);
            Assert.That(doc.FirstSection.Body.FirstParagraph.BreakIsStyleSeparator, Is.False);
        }

        [Test]
        public void TabStops()
        {
            //ExStart
            //ExFor:TabLeader
            //ExFor:TabAlignment
            //ExFor:Paragraph.GetEffectiveTabStops
            //ExSummary:Shows how to set custom tab stops for a paragraph.
            Document doc = new Document();
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // If we are in a paragraph with no tab stops in this collection,
            // the cursor will jump 36 points each time we press the Tab key in Microsoft Word.
            Assert.That(doc.FirstSection.Body.FirstParagraph.GetEffectiveTabStops().Length, Is.EqualTo(0));

            // We can add custom tab stops in Microsoft Word if we enable the ruler via the "View" tab.
            // Each unit on this ruler is two default tab stops, which is 72 points.
            // We can add custom tab stops programmatically like this.
            TabStopCollection tabStops = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.TabStops;
            tabStops.Add(72, TabAlignment.Left, TabLeader.Dots);
            tabStops.Add(216, TabAlignment.Center, TabLeader.Dashes);
            tabStops.Add(360, TabAlignment.Right, TabLeader.Line);

            // We can see these tab stops in Microsoft Word by enabling the ruler via "View" -> "Show" -> "Ruler".
            Assert.That(para.GetEffectiveTabStops().Length, Is.EqualTo(3));

            // Any tab characters we add will make use of the tab stops on the ruler and may,
            // depending on the tab leader's value, leave a line between the tab departure and arrival destinations.
            para.AppendChild(new Run(doc, "\tTab 1\tTab 2\tTab 3"));

            doc.Save(ArtifactsDir + "Paragraph.TabStops.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Paragraph.TabStops.docx");
            tabStops = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.TabStops;

            TestUtil.VerifyTabStop(72.0d, TabAlignment.Left, TabLeader.Dots, false, tabStops[0]);
            TestUtil.VerifyTabStop(216.0d, TabAlignment.Center, TabLeader.Dashes, false, tabStops[1]);
            TestUtil.VerifyTabStop(360.0d, TabAlignment.Right, TabLeader.Line, false, tabStops[2]);
        }

        [Test]
        public void JoinRuns()
        {
            //ExStart
            //ExFor:Paragraph.JoinRunsWithSameFormatting
            //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert four runs of text into the paragraph.
            builder.Write("Run 1. ");
            builder.Write("Run 2. ");
            builder.Write("Run 3. ");
            builder.Write("Run 4. ");

            // If we open this document in Microsoft Word, the paragraph will look like one seamless text body.
            // However, it will consist of four separate runs with the same formatting. Fragmented paragraphs like this
            // may occur when we manually edit parts of one paragraph many times in Microsoft Word.
            Paragraph para = builder.CurrentParagraph;

            Assert.That(para.Runs.Count, Is.EqualTo(4));

            // Change the style of the last run to set it apart from the first three.
            para.Runs[3].Font.StyleIdentifier = StyleIdentifier.Emphasis;

            // We can run the "JoinRunsWithSameFormatting" method to optimize the document's contents
            // by merging similar runs into one, reducing their overall count.
            // This method also returns the number of runs that this method merged.
            // These two merges occurred to combine Runs #1, #2, and #3,
            // while leaving out Run #4 because it has an incompatible style.
            Assert.That(para.JoinRunsWithSameFormatting(), Is.EqualTo(2));

            // The number of runs left will equal the original count
            // minus the number of run merges that the "JoinRunsWithSameFormatting" method carried out.
            Assert.That(para.Runs.Count, Is.EqualTo(2));
            Assert.That(para.Runs[0].Text, Is.EqualTo("Run 1. Run 2. Run 3. "));
            Assert.That(para.Runs[1].Text, Is.EqualTo("Run 4. "));
            //ExEnd
        }
    }
}