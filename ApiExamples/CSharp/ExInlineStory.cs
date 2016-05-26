// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;

using Aspose.Words;

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
            //ExFor:InlineStory
            //ExFor:InlineStory.Paragraphs
            //ExFor:InlineStory.FirstParagraph
            //ExFor:FootnoteType
            //ExFor:Footnote.#ctor
            //ExSummary:Shows how to add a footnote to a paragraph in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text is added.");

            Footnote footnote = new Footnote(doc, FootnoteType.Footnote);
            builder.CurrentParagraph.AppendChild(footnote);
            footnote.Paragraphs.Add(new Paragraph(doc));
            footnote.FirstParagraph.Runs.Add(new Run(doc, "Footnote text."));
            //ExEnd

            Assert.AreEqual("Footnote text.", doc.GetChildNodes(NodeType.Footnote, true)[0].ToString(SaveFormat.Text).Trim());
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
    }

}
