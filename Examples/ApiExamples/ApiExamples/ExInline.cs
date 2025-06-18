// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to determine the revision type of an inline node.
            Document doc = new Document(MyDir + "Revision runs.docx");

            // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
            // is turned on in Microsoft Word, the changes we apply count as revisions.
            // When editing a document using Aspose.Words, we can begin tracking revisions by
            // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
            // We can either accept revisions to assimilate them into the document
            // or reject them to change the proposed change effectively.
            Assert.That(doc.Revisions.Count, Is.EqualTo(6));

            // The parent node of a revision is the run that the revision concerns. A Run is an Inline node.
            Run run = (Run)doc.Revisions[0].ParentNode;

            Paragraph firstParagraph = run.ParentParagraph;
            RunCollection runs = firstParagraph.Runs;

            Assert.That(runs.ToArray().Length, Is.EqualTo(6));

            // Below are five types of revisions that can flag an Inline node.
            // 1 -  An "insert" revision:
            // This revision occurs when we insert text while tracking changes.
            Assert.That(runs[2].IsInsertRevision, Is.True);

            // 2 -  A "format" revision:
            // This revision occurs when we change the formatting of text while tracking changes.
            Assert.That(runs[2].IsFormatRevision, Is.True);

            // 3 -  A "move from" revision:
            // When we highlight text in Microsoft Word, and then drag it to a different place in the document
            // while tracking changes, two revisions appear.
            // The "move from" revision is a copy of the text originally before we moved it.
            Assert.That(runs[4].IsMoveFromRevision, Is.True);

            // 4 -  A "move to" revision:
            // The "move to" revision is the text that we moved in its new position in the document.
            // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
            // Accepting a move revision deletes the "move from" revision and its text,
            // and keeps the text from the "move to" revision.
            // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
            Assert.That(runs[1].IsMoveToRevision, Is.True);

            // 5 -  A "delete" revision:
            // This revision occurs when we delete text while tracking changes. When we delete text like this,
            // it will stay in the document as a revision until we either accept the revision,
            // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
            Assert.That(runs[5].IsDeleteRevision, Is.True);
            //ExEnd
        }
    }
}
