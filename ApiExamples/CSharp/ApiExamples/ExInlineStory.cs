// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExInlineStory : ApiExampleBase
    {
        [Test]
        public void AddFootnote()
        {
            //ExStart
            //ExFor:Footnote
            //ExFor:Footnote.IsAuto
            //ExFor:Footnote.ReferenceMark
            //ExFor:InlineStory
            //ExFor:InlineStory.Paragraphs
            //ExFor:InlineStory.FirstParagraph
            //ExFor:FootnoteType
            //ExFor:Footnote.#ctor
            //ExSummary:Shows how to add a footnote to a paragraph in the document and set its marker.
            // Create a new document and append some text that we will reference with a footnote
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Main body text.");

            // Add a footnote and give it text, which will appear at the bottom of the page
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // This attribute means the footnote in the main text will automatically be assigned a number, "1" in this instance
            // The next footnote will get "2"
            Assert.True(footnote.IsAuto);

            // We can edit the footnote's text like this
            // Make sure to move the builder back into the document body afterwards
            builder.MoveTo(footnote.FirstParagraph);
            builder.Write(" More text added by a DocumentBuilder.");
            builder.MoveToDocumentEnd();

            Assert.AreEqual("Footnote text. More text added by a DocumentBuilder.", footnote.Paragraphs[0].ToString(SaveFormat.Text).Trim());

            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // Substitute the reference number for our own custom mark by setting this variable, "IsAuto" will also be set to false
            footnote.ReferenceMark = "RefMark";
            Assert.False(footnote.IsAuto);

            // This bookmark will get a number "3" even though there was no "2"
            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");
            Assert.True(footnote.IsAuto);

            doc.Save(ArtifactsDir + "InlineStory.AddFootnote.docx");
            //ExEnd

            Assert.AreEqual("Footnote text. More text added by a DocumentBuilder.",
                doc.GetChildNodes(NodeType.Footnote, true)[0].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void FootnoteEndnote()
        {
            //ExStart
            //ExFor:Footnote.FootnoteType
            //ExSummary:Demonstrates the difference between footnotes and endnotes.
            // Create a document and a corresponding document builder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write text and insert a footnote to reference it at the bottom of the page
            builder.Write("Footnote referenced main body text.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text, will appear at the bottom of the page that contains the referenced text.");

            // Write text and insert an endnote to reference it at the end of the document
            builder.Write("Endnote referenced main body text.");
            Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, "Endnote text, will appear at the very end of the document.");

            // Since endnotes are at the end of the document, breaks like this will push them down while the footnotes stay where they are
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.AreEqual(FootnoteType.Footnote, footnote.FootnoteType);
            Assert.AreEqual(FootnoteType.Endnote, endnote.FootnoteType);

            doc.Save(ArtifactsDir + "InlineStory.FootnoteEndnote.docx");
            //ExEnd
        }

        [Test]
        public void AddComment()
        {
            //ExStart
            //ExFor:Comment
            //ExFor:InlineStory
            //ExFor:InlineStory.Paragraphs
            //ExFor:InlineStory.FirstParagraph
            //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
            //ExSummary:Shows how to add a comment to a paragraph in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text is added.");

            Comment comment = new Comment(doc, "Amy Lee", "AL", DateTime.Today);
            builder.CurrentParagraph.AppendChild(comment);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));
            //ExEnd

            Assert.AreEqual("Comment text.\r", (doc.GetChildNodes(NodeType.Comment, true)[0]).GetText());
        }

        [Test]
        public void InlineStoryRevisions()
        {
            //ExStart
            //ExFor:InlineStory.IsDeleteRevision
            //ExFor:InlineStory.IsInsertRevision
            //ExFor:InlineStory.IsMoveFromRevision
            //ExFor:InlineStory.IsMoveToRevision
            //ExSummary:Shows how to process revision-related properties of InlineStory nodes.
            // Open a document that has revisions from changes being tracked
            Document doc = new Document(MyDir + "InlineStory.Revisions.docx");
            Assert.IsTrue(doc.HasRevisions);

            // Get a collection of all footnotes from the document
            List<Footnote> footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();
            Assert.AreEqual(5, footnotes.Count);

            // If a node was inserted in Microsoft Word while changes were being tracked, this flag will be set to true
            Assert.IsTrue(footnotes[2].IsInsertRevision);

            // If one node was moved from one place to another while changes were tracked,
            // the node will be placed at the departure location as a "move to revision",
            // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
            // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
            // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
            Assert.IsTrue(footnotes[1].IsMoveToRevision);
            Assert.IsTrue(footnotes[4].IsMoveFromRevision);

            // If a node was deleted while changes were being tracked, it will stay behind as a delete revision until we accept/reject changes
            Assert.IsTrue(footnotes[3].IsDeleteRevision);
            //ExEnd
        }

        [Test]
        public void InsertInlineStoryNodes()
        {
            //ExStart
            //ExFor:Comment.StoryType
            //ExFor:Footnote.StoryType
            //ExFor:InlineStory.EnsureMinimum
            //ExFor:InlineStory.Font
            //ExFor:InlineStory.LastParagraph
            //ExFor:InlineStory.ParentParagraph
            //ExFor:InlineStory.StoryType
            //ExFor:InlineStory.Tables
            //ExSummary:Shows how to insert InlineStory nodes.
            // Create a new document and insert a blank footnote
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, null);

            // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell
            Table table = new Table(doc);
            table.EnsureMinimum();

            // We can place a table inside a footnote, which will make it appear at the footer of the referencing page
            Assert.That(footnote.Tables, Is.Empty);
            footnote.AppendChild(table);
            Assert.AreEqual(1, footnote.Tables.Count);
            Assert.AreEqual(NodeType.Table, footnote.LastChild.NodeType);

            // An InlineStory has an "EnsureMinimum()" method as well, but in this case it makes sure the last child of the node is a paragraph,
            // so we can click and write text easily in Microsoft Word
            footnote.EnsureMinimum();
            Assert.AreEqual(NodeType.Paragraph, footnote.LastChild.NodeType);

            // Edit the appearance of the anchor, which is the small superscript number in the main text that points to the footnote
            footnote.Font.Name = "Arial";
            footnote.Font.Color = Color.Green;

            // All inline story nodes have their own respective story types
            Assert.AreEqual(StoryType.Footnotes, footnote.StoryType);

            // A comment is another type of inline story
            Comment comment = (Comment)builder.CurrentParagraph.AppendChild(new Comment(doc, "John Doe", "J. D.", DateTime.Now));

            // The parent paragraph of an inline story node will be the one from the main document body
            Assert.AreEqual(doc.FirstSection.Body.FirstParagraph, comment.ParentParagraph);

            // However, the last paragraph is the one from the comment text contents, which will be outside the main document body in a speech bubble
            // A comment won't have any child nodes by default, so we can apply the EnsureMinimum() method to place a paragraph here as well
            Assert.Null(comment.LastParagraph);
            comment.EnsureMinimum();
            Assert.AreEqual(NodeType.Paragraph, comment.LastChild.NodeType);

            // Once we have a paragraph, we can move the builder do it and write our comment
            builder.MoveTo(comment.LastParagraph);
            builder.Write("My comment");

            Assert.AreEqual(StoryType.Comments, comment.StoryType);

            doc.Save(ArtifactsDir + "InlineStory.InsertInlineStoryNodes.docx");
            //ExEnd
        }
    }
}