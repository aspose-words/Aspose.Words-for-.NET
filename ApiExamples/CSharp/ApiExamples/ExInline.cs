// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
    class ExInline : ApiExampleBase
    {
        [Test]
        public void Inline()
        {
            //ExStart
            //ExFor:Inline
            //ExFor:Inline.IsDeleteRevision
            //ExFor:Inline.IsFormatRevision
            //ExFor:Inline.IsInsertRevision
            //ExFor:Inline.IsMoveFromRevision
            //ExFor:Inline.IsMoveToRevision
            //ExFor:Inline.ParentParagraph
            //ExSummary:Shows how to view revision-related properties of Inline nodes.
            Document doc = new Document(MyDir + "Inline.Revisions.docx");

            // This document has 5 revisions
            Assert.AreEqual(5, doc.Revisions.Count);

            // The parent node of a revision is the run that the revision concerns, which is an Inline node
            Run run = (Run)doc.Revisions[0].ParentNode;

            // Get the parent paragraph
            Paragraph firstParagraph = run.ParentParagraph;
            RunCollection runs = firstParagraph.Runs;

            Assert.AreEqual(6, runs.Count);

            // For all runs not involved in revisions, all the Is...Revision flags will be false

            // The text in the run at index #2 was typed after revisions were tracked, so it will count as an insert revision
            // The font was changed, so it will also be a format revision
            Assert.IsTrue(runs[2].IsInsertRevision);
            Assert.IsTrue(runs[2].IsFormatRevision);

            // For a "move", in regards to tracked revisions, to take place,
            // some text which contains at least one complete sentence must be removed from one location and placed into another
            // Typically this will happen when we highlight text with the mouse and drag it around, or cut and paste (but not copy and paste)
            // The node with the "IsMoveToRevision" flag is the destination, and the node with the "IsMoveFromRevision" flag is the departure point
            Assert.IsTrue(runs[1].IsMoveToRevision);
            Assert.IsTrue(runs[4].IsMoveFromRevision);
        }
    }
}
