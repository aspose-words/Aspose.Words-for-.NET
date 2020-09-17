// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExInlineStory : ApiExampleBase
    {
        [Test]
        public void Footnotes()
        {
            //ExStart
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.NumberStyle
            //ExFor:FootnoteOptions.Position
            //ExFor:FootnoteOptions.RestartRule
            //ExFor:FootnoteOptions.StartNumber
            //ExFor:FootnoteNumberingRule
            //ExFor:FootnotePosition
            //ExSummary:Shows how to insert footnotes, and modify their appearance.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3", "Custom reference mark");

            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            doc.FootnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;
            doc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            doc.FootnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "InlineStory.Footnotes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.Footnotes.docx");

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 1", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 2", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, false, "Custom reference mark",
                "Custom reference mark Footnote 3", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
        }

        [Test]
        public void Endnotes()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.NumberStyle
            //ExFor:EndnoteOptions.Position
            //ExFor:EndnoteOptions.RestartRule
            //ExFor:EndnoteOptions.StartNumber
            //ExFor:EndnotePosition
            //ExSummary:Shows how to insert endnotes, and modify their appearance.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 3", "Custom reference mark");
            
            doc.EndnoteOptions.Position = EndnotePosition.EndOfDocument;
            doc.EndnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;
            doc.EndnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            doc.EndnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "InlineStory.Endnotes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.Endnotes.docx");

            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 1", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 2", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, false, "Custom reference mark",
                "Custom reference mark Endnote 3", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
        }

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
            //ExSummary:Shows how to insert and customize footnotes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add text, and reference it with a footnote. This footnote will place a
            // small superscript reference mark after the text that it references, 
            // and will create an entry below the main body text at the bottom of the page.
            // This entry will contain the footnote's reference mark, as well as the reference text,
            // which we will pass to the document builder's "InsertFootnote" method.
            builder.Write("Main body text.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // If this property is set to "true", then our footnote's reference mark
            // will be its index among all of the section's footnotes.
            // This is the first footnote, so the reference mark will be "1".
            Assert.True(footnote.IsAuto);

            // We can move the document builder inside the footnote to edit its reference text. 
            builder.MoveTo(footnote.FirstParagraph);
            builder.Write(" More text added by a DocumentBuilder.");
            builder.MoveToDocumentEnd();

            Assert.AreEqual("\u0002 Footnote text. More text added by a DocumentBuilder.", footnote.GetText().Trim());

            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // We can set a custom reference mark which the footnote will use instead of its index number.
            footnote.ReferenceMark = "RefMark";

            Assert.False(footnote.IsAuto);

            // A bookmark with the "IsAuto" flag set to true will still show its real index
            // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            Assert.True(footnote.IsAuto);

            doc.Save(ArtifactsDir + "InlineStory.AddFootnote.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.AddFootnote.docx");

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, 
                "Footnote text. More text added by a DocumentBuilder.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, false, "RefMark", 
                "Footnote text.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, 
                "Footnote text.", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
        }

        [Test]
        public void FootnoteEndnote()
        {
            //ExStart
            //ExFor:Footnote.FootnoteType
            //ExSummary:Shows the difference between footnotes and endnotes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways of attaching numbered references to text. Both these types of references
            // will add a small superscript reference mark at the location that we insert them.
            // The reference mark, by default, is the index number of the reference among all the references in the document.
            // Each reference will also create an entry, which will have the same reference mark as in the body text,
            // as well as reference text, which we will pass to the document builder's "InsertFootnote" method.
            // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
            builder.Write("Footnote referenced main body text.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, 
                "Footnote text, will appear at the bottom of the page that contains the referenced text.");

            // 2 -  An endnote, whose entry will will appear at the end of the document:
            builder.Write("Endnote referenced main body text.");
            Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, 
                "Endnote text, will appear at the very end of the document.");

            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.AreEqual(FootnoteType.Footnote, footnote.FootnoteType);
            Assert.AreEqual(FootnoteType.Endnote, endnote.FootnoteType);

            doc.Save(ArtifactsDir + "InlineStory.FootnoteEndnote.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.FootnoteEndnote.docx");

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote text, will appear at the bottom of the page that contains the referenced text.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote text, will appear at the very end of the document.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
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
            //ExSummary:Shows how to add a comment to a paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Today);
            builder.CurrentParagraph.AppendChild(comment);
            builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
            builder.Write("Comment text.");

            // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it. 
            doc.Save(ArtifactsDir + "InlineStory.AddComment.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.AddComment.docx");
            comment = (Comment)doc.GetChild(NodeType.Comment, 0, true);
            
            Assert.AreEqual("Comment text.\r", comment.GetText());
            Assert.AreEqual("John Doe", comment.Author);
            Assert.AreEqual("JD", comment.Initial);
            Assert.AreEqual(DateTime.Today, comment.DateTime);
        }

        [Test]
        public void InlineStoryRevisions()
        {
            //ExStart
            //ExFor:InlineStory.IsDeleteRevision
            //ExFor:InlineStory.IsInsertRevision
            //ExFor:InlineStory.IsMoveFromRevision
            //ExFor:InlineStory.IsMoveToRevision
            //ExSummary:Shows how to view revision-related properties of InlineStory nodes.
            Document doc = new Document(MyDir + "Revision footnotes.docx");

            // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
            // is turned on in Microsoft Word, the changes we apply count as revisions.
            // When editing a document using Aspose.Words, we can begin tracking revisions by
            // invoking the document's "StartTrackRevisions" method, and stop tracking by using the "StopTrackRevisions" method.
            // We can either accept revisions to assimilate them into the document,
            // or reject them to undo and discard the change that they proposed.
            Assert.IsTrue(doc.HasRevisions);

            List<Footnote> footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();

            Assert.AreEqual(5, footnotes.Count);

            // Below are five types of revisions that an InlineStory node can be flagged as.
            // 1 -  An "insert" revision:
            // This revision occurs when we insert text while tracking changes.
            Assert.IsTrue(footnotes[2].IsInsertRevision);

            // 2 -  A "move from" revision:
            // When we highlight text in Microsoft Word, and then drag it to a different place in the document
            // while tracking changes, two revisions appear.
            // The "move from" revision is the copy of the text where it originally was before we moved it.
            Assert.IsTrue(footnotes[4].IsMoveFromRevision);

            // 3 -  A "move to" revision:
            // The "move to" revision is the text that we moved, in its new position in the document.
            // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
            // Accepting a move revision deletes the "move from" revision and its text,
            // and keeps the text from the "move to" revision.
            // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
            Assert.IsTrue(footnotes[1].IsMoveToRevision);

            // 4 -  A "delete" revision:
            // This revision occurs when we delete text while tracking changes. When we delete text like this,
            // it will stay in the document as a revision until we either accept the revision,
            // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, null);

            // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell.
            Table table = new Table(doc);
            table.EnsureMinimum();

            // We can place a table inside a footnote, which will make it appear at the footer of the referencing page
            Assert.That(footnote.Tables, Is.Empty);
            footnote.AppendChild(table);
            Assert.AreEqual(1, footnote.Tables.Count);
            Assert.AreEqual(NodeType.Table, footnote.LastChild.NodeType);

            // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
            // it makes sure the last child of the node is a paragraph, in order for us to be able to click and write text easily in Microsoft Word
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
            builder.Write("My comment.");

            Assert.AreEqual(StoryType.Comments, comment.StoryType);

            doc.Save(ArtifactsDir + "InlineStory.InsertInlineStoryNodes.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "InlineStory.InsertInlineStoryNodes.docx");

            footnote = (Footnote)doc.GetChild(NodeType.Footnote, 0, true);

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, string.Empty, 
                (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            Assert.AreEqual("Arial", footnote.Font.Name);
            Assert.AreEqual(Color.Green.ToArgb(), footnote.Font.Color.ToArgb());

            comment = (Comment)doc.GetChild(NodeType.Comment, 0, true);

            Assert.AreEqual("My comment.", comment.ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void DeleteShapes()
        {
            //ExStart
            //ExFor:Story
            //ExFor:Story.DeleteShapes
            //ExFor:Story.StoryType
            //ExFor:StoryType
            //ExSummary:Shows how to remove all shapes from a node.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert a shape. This is an inline shape,
            // which has a parent Paragraph, which is a child node of the first section's Body.
            builder.InsertShape(ShapeType.Cube, 100.0, 100.0);

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);

            // We can delete all shapes from the child paragraphs of this Body.
            Assert.AreEqual(StoryType.MainText, doc.FirstSection.Body.StoryType);
            doc.FirstSection.Body.DeleteShapes();

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);
            //ExEnd
        }
    }
}