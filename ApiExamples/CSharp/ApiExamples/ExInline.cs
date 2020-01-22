// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExInline : ApiExampleBase
    {
        [Test]
        public void InlineRevisions()
        {
            //ExStart
            //ExFor:Inline
            //ExFor:Inline.IsDeleteRevision
            //ExFor:Inline.IsFormatRevision
            //ExFor:Inline.IsInsertRevision
            //ExFor:Inline.IsMoveFromRevision
            //ExFor:Inline.IsMoveToRevision
            //ExFor:Inline.ParentParagraph
            //ExFor:Paragraph.Runs
            //ExFor:Revision.ParentNode
            //ExFor:RunCollection
            //ExFor:RunCollection.Item(Int32)
            //ExFor:RunCollection.ToArray
            //ExSummary:Shows how to process revision-related properties of Inline nodes.
            Document doc = new Document(MyDir + "Inline.Revisions.docx");

            // This document has 6 revisions
            Assert.AreEqual(6, doc.Revisions.Count);

            // The parent node of a revision is the run that the revision concerns, which is an Inline node
            Run run = (Run)doc.Revisions[0].ParentNode;

            // Get the parent paragraph
            Paragraph firstParagraph = run.ParentParagraph;
            RunCollection runs = firstParagraph.Runs;

            Assert.AreEqual(6, runs.ToArray().Length);

            // The text in the run at index #2 was typed after revisions were tracked, so it will count as an insert revision
            // The font was changed, so it will also be a format revision
            Assert.IsTrue(runs[2].IsInsertRevision);
            Assert.IsTrue(runs[2].IsFormatRevision);

            // If one node was moved from one place to another while changes were tracked,
            // the node will be placed at the departure location as a "move to revision",
            // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
            // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
            // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
            Assert.IsTrue(runs[1].IsMoveToRevision);
            Assert.IsTrue(runs[4].IsMoveFromRevision);

            // If an Inline node gets deleted while changes are being tracked, it will leave behind a node with the IsDeleteRevision flag set to true until changes are accepted
            Assert.IsTrue(runs[5].IsDeleteRevision);
            //ExEnd
        }
    }
}
