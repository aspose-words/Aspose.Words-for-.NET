// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExRevision : ApiExampleBase
    {
        [Test]
        public void Revisions()
        {
            //ExStart
            //ExFor:Revision
            //ExFor:Revision.Accept
            //ExFor:Revision.Author
            //ExFor:Revision.DateTime
            //ExFor:Revision.Group
            //ExFor:Revision.Reject
            //ExFor:Revision.RevisionType
            //ExFor:RevisionCollection
            //ExFor:RevisionCollection.Item(Int32)
            //ExFor:RevisionCollection.Count
            //ExFor:RevisionType
            //ExFor:Document.HasRevisions
            //ExFor:Document.TrackRevisions
            //ExFor:Document.Revisions
            //ExSummary:Shows how to work with revisions in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normal editing of the document does not count as a revision.
            builder.Write("This does not count as a revision. ");

            Assert.IsFalse(doc.HasRevisions);

            // To register our edits as revisions, we need to declare an author, and then start tracking them.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            builder.Write("This is revision #1. ");

            Assert.IsTrue(doc.HasRevisions);
            Assert.AreEqual(1, doc.Revisions.Count);

            // This flag corresponds to the Review -> Tracking -> "Track Changes" option being turned on in Microsoft Word, 
            // and it is independent of the programmatic revision tracking that is taking place here.
            Assert.IsFalse(doc.TrackRevisions);

            // Our first revision is an insertion-type revision, since we added text with the document builder.
            Revision revision = doc.Revisions[0];
            Assert.AreEqual("John Doe", revision.Author);
            Assert.AreEqual("This is revision #1. ", revision.ParentNode.GetText());
            Assert.AreEqual(RevisionType.Insertion, revision.RevisionType);
            Assert.AreEqual(revision.DateTime.Date, DateTime.Now.Date);
            Assert.AreEqual(doc.Revisions.Groups[0], revision.Group);

            // Remove a run to create a deletion-type revision.
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

            // Every new revision is put at the beginning of the revision collection.
            Assert.AreEqual(RevisionType.Deletion, doc.Revisions[0].RevisionType);
            Assert.AreEqual(2, doc.Revisions.Count);

            // Insert revisions are treated as document text by the GetText() method before they are accepted,
            // since they are still nodes with text and are in the body.
            Assert.AreEqual("This does not count as a revision. This is revision #1.", doc.GetText().Trim());

            // Accepting the delete revision will remove its parent node from the paragraph text,
            // and then remove the revision itself from the collection.
            doc.Revisions[0].Accept();

            Assert.AreEqual(1, doc.Revisions.Count);

            // Once the delete revision is accepted, the nodes that it concerns are removed,
            // and their contents will no longer be anywhere in the document.
            Assert.AreEqual("This is revision #1.", doc.GetText().Trim());

            // The insertion-type revision is now at index 0, which we can reject to ignore and discard it
            doc.Revisions[0].Reject();

            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.AreEqual("", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void RevisionCollection()
        {
            //ExStart
            //ExFor:Revision.ParentStyle
            //ExFor:RevisionCollection.GetEnumerator
            //ExFor:RevisionCollection.Groups
            //ExFor:RevisionCollection.RejectAll
            //ExFor:RevisionGroupCollection.GetEnumerator
            //ExSummary:Shows how to iterate through a document's revisions.
            Document doc = new Document(MyDir + "Revisions.docx");
            RevisionCollection revisions = doc.Revisions;

            // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions.
            Assert.AreEqual(7, revisions.Groups.Count); //ExSkip
            Console.WriteLine($"{revisions.Groups.Count} revision groups:");

            // Iterate over the collection of groups and print the text that the revision concerns.
            using (IEnumerator<RevisionGroup> e = revisions.Groups.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    Console.WriteLine($"\tGroup type \"{e.Current.RevisionType}\", " +
                                      $"author: {e.Current.Author}, contents: [{e.Current.Text.Trim()}]");
                }
            }

            // Each Run affected by a revision gets its own Revision object.
            // The collection of revisions is considerably larger than the condensed form we printed above,
            // depending on how many Runs the text has been segmented into during editing in Microsoft Word.
            Assert.AreEqual(11, revisions.Count); //ExSkip
            Console.WriteLine($"\n{revisions.Count} revisions:");

            using (IEnumerator<Revision> e = revisions.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    // A StyleDefinitionChange strictly affects styles and not document nodes, so in this case the ParentStyle
                    // attribute will always be used, while the ParentNode will always be null.
                    // Since all other changes affect nodes, ParentNode will conversely be in use and ParentStyle will be null.
                    if (e.Current.RevisionType == RevisionType.StyleDefinitionChange)
                    {
                        Console.WriteLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, style: [{e.Current.ParentStyle.Name}]");
                    }
                    else
                    {
                        Console.WriteLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, contents: [{e.Current.ParentNode.GetText().Trim()}]");
                    }
                }
            }

            // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
            // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document.
            // In this case, we will reject all revisions via the collection, reverting the document to its original form.
            revisions.RejectAll();

            Assert.AreEqual(0, revisions.Count);
            //ExEnd
        }
    }
}
