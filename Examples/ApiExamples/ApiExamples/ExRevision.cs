// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Layout;
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

            // This flag corresponds to the "Review" -> "Tracking" -> "Track Changes" option in Microsoft Word.
            // The "StartTrackRevisions" method does not affect its value,
            // and the document is tracking revisions programmatically despite it having a value of "false".
            // If we open this document using Microsoft Word, it will not be tracking revisions.
            Assert.IsFalse(doc.TrackRevisions);

            // We have added text using the document builder, so the first revision is an insertion-type revision.
            Revision revision = doc.Revisions[0];
            Assert.AreEqual("John Doe", revision.Author);
            Assert.AreEqual("This is revision #1. ", revision.ParentNode.GetText());
            Assert.AreEqual(RevisionType.Insertion, revision.RevisionType);
            Assert.AreEqual(revision.DateTime.Date, DateTime.Now.Date);
            Assert.AreEqual(doc.Revisions.Groups[0], revision.Group);

            // Remove a run to create a deletion-type revision.
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

            // Adding a new revision places it at the beginning of the revision collection.
            Assert.AreEqual(RevisionType.Deletion, doc.Revisions[0].RevisionType);
            Assert.AreEqual(2, doc.Revisions.Count);

            // Insert revisions show up in the document body even before we accept/reject the revision.
            // Rejecting the revision will remove its nodes from the body. Conversely, nodes that make up delete revisions
            // also linger in the document until we accept the revision.
            Assert.AreEqual("This does not count as a revision. This is revision #1.", doc.GetText().Trim());

            // Accepting the delete revision will remove its parent node from the paragraph text
            // and then remove the collection's revision itself.
            doc.Revisions[0].Accept();

            Assert.AreEqual(1, doc.Revisions.Count);
            Assert.AreEqual("This is revision #1.", doc.GetText().Trim());

            // The insertion-type revision is now at index 0. Reject the revision to discard its contents.
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
            //ExSummary:Shows how to work with a document's collection of revisions.
            Document doc = new Document(MyDir + "Revisions.docx");
            RevisionCollection revisions = doc.Revisions;

            // This collection itself has a collection of revision groups.
            // Each group is a sequence of adjacent revisions.
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

            // Each Run that a revision affects gets a corresponding Revision object.
            // The revisions' collection is considerably larger than the condensed form we printed above,
            // depending on how many Runs we have segmented the document into during Microsoft Word editing.
            Assert.AreEqual(11, revisions.Count); //ExSkip
            Console.WriteLine($"\n{revisions.Count} revisions:");

            using (IEnumerator<Revision> e = revisions.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    // A StyleDefinitionChange strictly affects styles and not document nodes. This means the "ParentStyle"
                    // property will always be in use, while the ParentNode will always be null.
                    // Since all other changes affect nodes, ParentNode will conversely be in use, and ParentStyle will be null.
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

            // Reject all revisions via the collection, reverting the document to its original form.
            revisions.RejectAll();

            Assert.AreEqual(0, revisions.Count);
            //ExEnd
        }

        [Test]
        public void GetInfoAboutRevisionsInRevisionGroups()
        {
            //ExStart
            //ExFor:RevisionGroup
            //ExFor:RevisionGroup.Author
            //ExFor:RevisionGroup.RevisionType
            //ExFor:RevisionGroup.Text
            //ExFor:RevisionGroupCollection
            //ExFor:RevisionGroupCollection.Count
            //ExSummary:Shows how to print info about a group of revisions in a document.
            Document doc = new Document(MyDir + "Revisions.docx");

            Assert.AreEqual(7, doc.Revisions.Groups.Count);

            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine(
                    $"Revision author: {group.Author}; Revision type: {group.RevisionType} \n\tRevision text: {group.Text}");
            }
            //ExEnd
        }

        [Test]
        public void GetSpecificRevisionGroup()
        {
            //ExStart
            //ExFor:RevisionGroupCollection
            //ExFor:RevisionGroupCollection.Item(Int32)
            //ExSummary:Shows how to get a group of revisions in a document.
            Document doc = new Document(MyDir + "Revisions.docx");

            RevisionGroup revisionGroup = doc.Revisions.Groups[0];
            //ExEnd

            Assert.AreEqual(RevisionType.Deletion, revisionGroup.RevisionType);
            Assert.AreEqual("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
                revisionGroup.Text);
        }

        [Test]
        public void ShowRevisionBalloons()
        {
            //ExStart
            //ExFor:RevisionOptions.ShowInBalloons
            //ExSummary:Shows how to display revisions in balloons.
            Document doc = new Document(MyDir + "Revisions.docx");

            // By default, text that is a revision has a different color to differentiate it from the other non-revision text.
            // Set a revision option to show more details about each revision in a balloon on the page's right margin.
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;
            doc.Save(ArtifactsDir + "Revision.ShowRevisionBalloons.pdf");
            //ExEnd
        }

        [Test]
        public void RevisionOptions()
        {
            //ExStart
            //ExFor:ShowInBalloons
            //ExFor:RevisionOptions.ShowInBalloons
            //ExFor:RevisionOptions.CommentColor
            //ExFor:RevisionOptions.DeletedTextColor
            //ExFor:RevisionOptions.DeletedTextEffect
            //ExFor:RevisionOptions.InsertedTextEffect
            //ExFor:RevisionOptions.MovedFromTextColor
            //ExFor:RevisionOptions.MovedFromTextEffect
            //ExFor:RevisionOptions.MovedToTextColor
            //ExFor:RevisionOptions.MovedToTextEffect
            //ExFor:RevisionOptions.RevisedPropertiesColor
            //ExFor:RevisionOptions.RevisedPropertiesEffect
            //ExFor:RevisionOptions.RevisionBarsColor
            //ExFor:RevisionOptions.RevisionBarsWidth
            //ExFor:RevisionOptions.ShowOriginalRevision
            //ExFor:RevisionOptions.ShowRevisionMarks
            //ExFor:RevisionTextEffect
            //ExSummary:Shows how to modify the appearance of revisions.
            Document doc = new Document(MyDir + "Revisions.docx");

            // Get the RevisionOptions object that controls the appearance of revisions.
            RevisionOptions revisionOptions = doc.LayoutOptions.RevisionOptions;

            // Render insertion revisions in green and italic.
            revisionOptions.InsertedTextColor = RevisionColor.Green;
            revisionOptions.InsertedTextEffect = RevisionTextEffect.Italic;

            // Render deletion revisions in red and bold.
            revisionOptions.DeletedTextColor = RevisionColor.Red;
            revisionOptions.DeletedTextEffect = RevisionTextEffect.Bold;

            // The same text will appear twice in a movement revision:
            // once at the departure point and once at the arrival destination.
            // Render the text at the moved-from revision yellow with a double strike through
            // and double-underlined blue at the moved-to revision.
            revisionOptions.MovedFromTextColor = RevisionColor.Yellow;
            revisionOptions.MovedFromTextEffect = RevisionTextEffect.DoubleStrikeThrough;
            revisionOptions.MovedToTextColor = RevisionColor.Blue;
            revisionOptions.MovedFromTextEffect = RevisionTextEffect.DoubleUnderline;

            // Render format revisions in dark red and bold.
            revisionOptions.RevisedPropertiesColor = RevisionColor.DarkRed;
            revisionOptions.RevisedPropertiesEffect = RevisionTextEffect.Bold;

            // Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
            revisionOptions.RevisionBarsColor = RevisionColor.DarkBlue;
            revisionOptions.RevisionBarsWidth = 15.0f;

            // Show revision marks and original text.
            revisionOptions.ShowOriginalRevision = true;
            revisionOptions.ShowRevisionMarks = true;

            // Get movement, deletion, formatting revisions, and comments to show up in green balloons
            // on the right side of the page.
            revisionOptions.ShowInBalloons = ShowInBalloons.Format;
            revisionOptions.CommentColor = RevisionColor.BrightGreen;

            // These features are only applicable to formats such as .pdf or .jpg.
            doc.Save(ArtifactsDir + "Revision.RevisionOptions.pdf");
            //ExEnd
        }
    }
}
