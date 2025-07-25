﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Notes;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExInlineStory : ApiExampleBase
    {
        [TestCase(FootnotePosition.BeneathText)]
        [TestCase(FootnotePosition.BottomOfPage)]
        public void PositionFootnote(FootnotePosition footnotePosition)
        {
            //ExStart
            //ExFor:Document.FootnoteOptions
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.Position
            //ExFor:FootnotePosition
            //ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A footnote is a way to attach a reference or a side comment to text
            // that does not interfere with the main body text's flow.  
            // Inserting a footnote adds a small superscript reference symbol
            // at the main body text where we insert the footnote.
            // Each footnote also creates an entry at the bottom of the page, consisting of a symbol
            // that matches the reference symbol in the main body text.
            // The reference text that we pass to the document builder's "InsertFootnote" method.
            builder.Write("Hello world!");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote contents.");

            // We can use the "Position" property to determine where the document will place all its footnotes.
            // If we set the value of the "Position" property to "FootnotePosition.BottomOfPage",
            // every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
            // If we set the value of the "Position" property to "FootnotePosition.BeneathText",
            // every footnote will show up at the end of the page's text that contains its reference mark.
            doc.FootnoteOptions.Position = footnotePosition;

            doc.Save(ArtifactsDir + "InlineStory.PositionFootnote.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.PositionFootnote.docx");

            Assert.That(doc.FootnoteOptions.Position, Is.EqualTo(footnotePosition));

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote contents.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
        }

        [TestCase(EndnotePosition.EndOfDocument)]
        [TestCase(EndnotePosition.EndOfSection)]
        public void PositionEndnote(EndnotePosition endnotePosition)
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.Position
            //ExFor:EndnotePosition
            //ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // An endnote is a way to attach a reference or a side comment to text
            // that does not interfere with the main body text's flow. 
            // Inserting an endnote adds a small superscript reference symbol
            // at the main body text where we insert the endnote.
            // Each endnote also creates an entry at the end of the document, consisting of a symbol
            // that matches the reference symbol in the main body text.
            // The reference text that we pass to the document builder's "InsertEndnote" method.
            builder.Write("Hello world!");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote contents.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("This is the second section.");

            // We can use the "Position" property to determine where the document will place all its endnotes.
            // If we set the value of the "Position" property to "EndnotePosition.EndOfDocument",
            // every footnote will show up in a collection at the end of the document. This is the default value.
            // If we set the value of the "Position" property to "EndnotePosition.EndOfSection",
            // every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
            doc.EndnoteOptions.Position = endnotePosition;

            doc.Save(ArtifactsDir + "InlineStory.PositionEndnote.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.PositionEndnote.docx");

            Assert.That(doc.EndnoteOptions.Position, Is.EqualTo(endnotePosition));

            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote contents.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
        }

        [Test]
        public void RefMarkNumberStyle()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.NumberStyle
            //ExFor:Document.FootnoteOptions
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.NumberStyle
            //ExSummary:Shows how to change the number style of footnote/endnote reference marks.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Footnotes and endnotes are a way to attach a reference or a side comment to text
            // that does not interfere with the main body text's flow. 
            // Inserting a footnote/endnote adds a small superscript reference symbol
            // at the main body text where we insert the footnote/endnote.
            // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
            // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
            // Footnote entries, by default, show up at the bottom of each page that contains
            // their reference symbols, and endnotes show up at the end of the document.
            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2.");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3.", "Custom footnote reference mark");

            builder.InsertParagraph();

            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2.");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 3.", "Custom endnote reference mark");

            // By default, the reference symbol for each footnote and endnote is its index
            // among all the document's footnotes/endnotes. Each document maintains separate counts
            // for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
            // and endnotes display their numbers in lowercase Roman numerals.
            Assert.That(doc.FootnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.Arabic));
            Assert.That(doc.EndnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.LowercaseRoman));

            // We can use the "NumberStyle" property to apply custom numbering styles to footnotes and endnotes.
            // This will not affect footnotes/endnotes with custom reference marks.
            doc.FootnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;
            doc.EndnoteOptions.NumberStyle = NumberStyle.UppercaseLetter;

            doc.Save(ArtifactsDir + "InlineStory.RefMarkNumberStyle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.RefMarkNumberStyle.docx");

            Assert.That(doc.FootnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.UppercaseRoman));
            Assert.That(doc.EndnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.UppercaseLetter));

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, false, "Custom footnote reference mark",
                "Custom footnote reference mark Footnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 3, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 4, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, false, "Custom endnote reference mark",
                "Custom endnote reference mark Endnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 5, true));
        }

        [Test]
        public void NumberingRule()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.RestartRule
            //ExFor:FootnoteNumberingRule
            //ExFor:Document.FootnoteOptions
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.RestartRule
            //ExSummary:Shows how to restart footnote/endnote numbering at certain places in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Footnotes and endnotes are a way to attach a reference or a side comment to text
            // that does not interfere with the main body text's flow. 
            // Inserting a footnote/endnote adds a small superscript reference symbol
            // at the main body text where we insert the footnote/endnote.
            // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
            // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
            // Footnote entries, by default, show up at the bottom of each page that contains
            // their reference symbols, and endnotes show up at the end of the document.
            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3.");
            builder.Write("Text 4. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 4.");

            builder.InsertBreak(BreakType.PageBreak);

            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 3.");
            builder.Write("Text 4. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 4.");

            // By default, the reference symbol for each footnote and endnote is its index
            // among all the document's footnotes/endnotes. Each document maintains separate counts
            // for footnotes and endnotes and does not restart these counts at any point.
            Assert.That(FootnoteNumberingRule.Default, Is.EqualTo(doc.FootnoteOptions.RestartRule));
            Assert.That(FootnoteNumberingRule.Continuous, Is.EqualTo(FootnoteNumberingRule.Default));

            // We can use the "RestartRule" property to get the document to restart
            // the footnote/endnote counts at a new page or section.
            doc.FootnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            doc.EndnoteOptions.RestartRule = FootnoteNumberingRule.RestartSection;

            doc.Save(ArtifactsDir + "InlineStory.NumberingRule.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.NumberingRule.docx");

            Assert.That(doc.FootnoteOptions.RestartRule, Is.EqualTo(FootnoteNumberingRule.RestartPage));
            Assert.That(doc.EndnoteOptions.RestartRule, Is.EqualTo(FootnoteNumberingRule.RestartSection));

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 4.", (Footnote)doc.GetChild(NodeType.Footnote, 3, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 4, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 5, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 6, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 4.", (Footnote)doc.GetChild(NodeType.Footnote, 7, true));
        }

        [Test]
        public void StartNumber()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.StartNumber
            //ExFor:Document.FootnoteOptions
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.StartNumber
            //ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Footnotes and endnotes are a way to attach a reference or a side comment to text
            // that does not interfere with the main body text's flow. 
            // Inserting a footnote/endnote adds a small superscript reference symbol
            // at the main body text where we insert the footnote/endnote.
            // Each footnote/endnote also creates an entry, which consists of a symbol
            // that matches the reference symbol in the main body text.
            // The reference text that we pass to the document builder's "InsertEndnote" method.
            // Footnote entries, by default, show up at the bottom of each page that contains
            // their reference symbols, and endnotes show up at the end of the document.
            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2.");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3.");

            builder.InsertParagraph();

            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1.");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2.");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 3.");

            // By default, the reference symbol for each footnote and endnote is its index
            // among all the document's footnotes/endnotes. Each document maintains separate counts
            // for footnotes and for endnotes, which both begin at 1.
            Assert.That(doc.FootnoteOptions.StartNumber, Is.EqualTo(1));
            Assert.That(doc.EndnoteOptions.StartNumber, Is.EqualTo(1));

            // We can use the "StartNumber" property to get the document to
            // begin a footnote or endnote count at a different number.
            doc.EndnoteOptions.NumberStyle = NumberStyle.Arabic;
            doc.EndnoteOptions.StartNumber = 50;

            doc.Save(ArtifactsDir + "InlineStory.StartNumber.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.StartNumber.docx");

            Assert.That(doc.FootnoteOptions.StartNumber, Is.EqualTo(1));
            Assert.That(doc.EndnoteOptions.StartNumber, Is.EqualTo(50));
            Assert.That(doc.FootnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.Arabic));
            Assert.That(doc.EndnoteOptions.NumberStyle, Is.EqualTo(NumberStyle.Arabic));

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 1.", (Footnote)doc.GetChild(NodeType.Footnote, 3, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 2.", (Footnote)doc.GetChild(NodeType.Footnote, 4, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 3.", (Footnote)doc.GetChild(NodeType.Footnote, 5, true));
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

            // Add text, and reference it with a footnote. This footnote will place a small superscript reference
            // mark after the text that it references and create an entry below the main body text at the bottom of the page.
            // This entry will contain the footnote's reference mark and the reference text,
            // which we will pass to the document builder's "InsertFootnote" method.
            builder.Write("Main body text.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // If this property is set to "true", then our footnote's reference mark
            // will be its index among all the section's footnotes.
            // This is the first footnote, so the reference mark will be "1".
            Assert.That(footnote.IsAuto, Is.True);

            // We can move the document builder inside the footnote to edit its reference text. 
            builder.MoveTo(footnote.FirstParagraph);
            builder.Write(" More text added by a DocumentBuilder.");
            builder.MoveToDocumentEnd();

            Assert.That(footnote.GetText().Trim(), Is.EqualTo("\u0002 Footnote text. More text added by a DocumentBuilder."));

            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            // We can set a custom reference mark which the footnote will use instead of its index number.
            footnote.ReferenceMark = "RefMark";

            Assert.That(footnote.IsAuto, Is.False);

            // A bookmark with the "IsAuto" flag set to true will still show its real index
            // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
            builder.Write(" More main body text.");
            footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

            Assert.That(footnote.IsAuto, Is.True);

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

            // Below are two ways of attaching numbered references to the text. Both these references will add a
            // small superscript reference mark at the location that we insert them.
            // The reference mark, by default, is the index number of the reference among all the references in the document.
            // Each reference will also create an entry, which will have the same reference mark as in the body text
            // and reference text, which we will pass to the document builder's "InsertFootnote" method.
            // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
            builder.Write("Footnote referenced main body text.");
            Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, 
                "Footnote text, will appear at the bottom of the page that contains the referenced text.");

            // 2 -  An endnote, whose entry will appear at the end of the document:
            builder.Write("Endnote referenced main body text.");
            Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, 
                "Endnote text, will appear at the very end of the document.");

            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.That(footnote.FootnoteType, Is.EqualTo(FootnoteType.Footnote));
            Assert.That(endnote.FootnoteType, Is.EqualTo(FootnoteType.Endnote));

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

            Assert.That(comment.DateTime, Is.EqualTo(DateTime.Today));

            // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it. 
            doc.Save(ArtifactsDir + "InlineStory.AddComment.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.AddComment.docx");
            comment = (Comment)doc.GetChild(NodeType.Comment, 0, true);
            
            Assert.That(comment.GetText(), Is.EqualTo("Comment text.\r"));
            Assert.That(comment.Author, Is.EqualTo("John Doe"));
            Assert.That(comment.Initial, Is.EqualTo("JD"));
            Assert.That(comment.DateTime, Is.EqualTo(DateTime.Today));
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
            // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
            // We can either accept revisions to assimilate them into the document
            // or reject them to undo and discard the proposed change.
            Assert.That(doc.HasRevisions, Is.True);

            List<Footnote> footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>().ToList();

            Assert.That(footnotes.Count, Is.EqualTo(5));

            // Below are five types of revisions that can flag an InlineStory node.
            // 1 -  An "insert" revision:
            // This revision occurs when we insert text while tracking changes.
            Assert.That(footnotes[2].IsInsertRevision, Is.True);

            // 2 -  A "move from" revision:
            // When we highlight text in Microsoft Word, and then drag it to a different place in the document
            // while tracking changes, two revisions appear.
            // The "move from" revision is a copy of the text originally before we moved it.
            Assert.That(footnotes[4].IsMoveFromRevision, Is.True);

            // 3 -  A "move to" revision:
            // The "move to" revision is the text that we moved in its new position in the document.
            // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
            // Accepting a move revision deletes the "move from" revision and its text,
            // and keeps the text from the "move to" revision.
            // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
            Assert.That(footnotes[1].IsMoveToRevision, Is.True);

            // 4 -  A "delete" revision:
            // This revision occurs when we delete text while tracking changes. When we delete text like this,
            // it will stay in the document as a revision until we either accept the revision,
            // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
            Assert.That(footnotes[3].IsDeleteRevision, Is.True);
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

            // We can place a table inside a footnote, which will make it appear at the referencing page's footer.
            Assert.That(footnote.Tables.Count, Is.EqualTo(0));
            footnote.AppendChild(table);
            Assert.That(footnote.Tables.Count, Is.EqualTo(1));
            Assert.That(footnote.LastChild.NodeType, Is.EqualTo(NodeType.Table));

            // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
            // it makes sure the last child of the node is a paragraph,
            // for us to be able to click and write text easily in Microsoft Word.
            footnote.EnsureMinimum();
            Assert.That(footnote.LastChild.NodeType, Is.EqualTo(NodeType.Paragraph));

            // Edit the appearance of the anchor, which is the small superscript number
            // in the main text that points to the footnote.
            footnote.Font.Name = "Arial";
            footnote.Font.Color = Color.Green;

            // All inline story nodes have their respective story types.
            Assert.That(footnote.StoryType, Is.EqualTo(StoryType.Footnotes));

            // A comment is another type of inline story.
            Comment comment = (Comment)builder.CurrentParagraph.AppendChild(new Comment(doc, "John Doe", "J. D.", DateTime.Now));

            // The parent paragraph of an inline story node will be the one from the main document body.
            Assert.That(comment.ParentParagraph, Is.EqualTo(doc.FirstSection.Body.FirstParagraph));

            // However, the last paragraph is the one from the comment text contents,
            // which will be outside the main document body in a speech bubble.
            // A comment will not have any child nodes by default,
            // so we can apply the EnsureMinimum() method to place a paragraph here as well.
            Assert.That(comment.LastParagraph, Is.Null);
            comment.EnsureMinimum();
            Assert.That(comment.LastChild.NodeType, Is.EqualTo(NodeType.Paragraph));

            // Once we have a paragraph, we can move the builder to do it and write our comment.
            builder.MoveTo(comment.LastParagraph);
            builder.Write("My comment.");

            Assert.That(comment.StoryType, Is.EqualTo(StoryType.Comments));

            doc.Save(ArtifactsDir + "InlineStory.InsertInlineStoryNodes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "InlineStory.InsertInlineStoryNodes.docx");

            footnote = (Footnote)doc.GetChild(NodeType.Footnote, 0, true);

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, string.Empty, 
                (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            Assert.That(footnote.Font.Name, Is.EqualTo("Arial"));
            Assert.That(footnote.Font.Color.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));

            comment = (Comment)doc.GetChild(NodeType.Comment, 0, true);

            Assert.That(comment.ToString(SaveFormat.Text).Trim(), Is.EqualTo("My comment."));
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

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(1));

            // We can delete all shapes from the child paragraphs of this Body.
            Assert.That(doc.FirstSection.Body.StoryType, Is.EqualTo(StoryType.MainText));
            doc.FirstSection.Body.DeleteShapes();

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void UpdateActualReferenceMarks()
        {
            //ExStart:UpdateActualReferenceMarks
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:Document.UpdateActualReferenceMarks
            //ExFor:Footnote.ActualReferenceMark
            //ExSummary:Shows how to get actual footnote reference mark.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");

            Footnote footnote = (Footnote)doc.GetChild(NodeType.Footnote, 1, true);
            doc.UpdateFields();
            doc.UpdateActualReferenceMarks();

            Assert.That(footnote.ActualReferenceMark, Is.EqualTo("1"));
            //ExEnd:UpdateActualReferenceMarks
        }

        [Test]
        public void EndnoteSeparator()
        {
            //ExStart:EndnoteSeparator
            //GistId:e06aa7a168b57907a5598e823a22bf0a
            //ExFor:DocumentBase.FootnoteSeparators
            //ExFor:FootnoteSeparatorType
            //ExSummary:Shows how to remove endnote separator.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");

            FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
            // Remove endnote separator.
            endnoteSeparator.FirstParagraph.FirstChild.Remove();
            //ExEnd:EndnoteSeparator

            doc.Save(ArtifactsDir + "InlineStory.EndnoteSeparator.docx");
        }

        [Test]
        public void FootnoteSeparator()
        {
            //ExStart:FootnoteSeparator
            //GistId:e06aa7a168b57907a5598e823a22bf0a
            //ExFor:DocumentBase.FootnoteSeparators
            //ExFor:FootnoteSeparator
            //ExFor:FootnoteSeparatorType
            //ExFor:FootnoteSeparatorCollection
            //ExFor:FootnoteSeparatorCollection.Item(FootnoteSeparatorType)
            //ExSummary:Shows how to manage footnote separator format.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");

            FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
            // Align footnote separator.
            footnoteSeparator.FirstParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            //ExEnd:FootnoteSeparator

            doc.Save(ArtifactsDir + "InlineStory.FootnoteSeparator.docx");
        }
    }
}
