// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Notes;
using Aspose.Words.Tables;
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

            builder.Writeln("");
            builder.Write("This is revision #2.");

            // Now move the node to create a moving revision type.
            Node node = doc.FirstSection.Body.Paragraphs[1];
            Node endNode = doc.FirstSection.Body.Paragraphs[1].NextSibling;
            Node referenceNode = doc.FirstSection.Body.Paragraphs[0];

            while (node != endNode)
            {
                Node nextNode = node.NextSibling;
                doc.FirstSection.Body.InsertBefore(node, referenceNode);
                node = nextNode;
            }

            Assert.AreEqual(RevisionType.Moving, doc.Revisions[0].RevisionType);
            Assert.AreEqual(8, doc.Revisions.Count);
            Assert.AreEqual("This is revision #2.\rThis is revision #1. \rThis is revision #2.", doc.GetText().Trim());

            // The moving revision is now at index 1. Reject the revision to discard its contents.
            doc.Revisions[1].Reject();

            Assert.AreEqual(6, doc.Revisions.Count);
            Assert.AreEqual("This is revision #1. \rThis is revision #2.", doc.GetText().Trim());
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
            revisionOptions.MovedToTextColor = RevisionColor.ClassicBlue;
            revisionOptions.MovedToTextEffect = RevisionTextEffect.DoubleUnderline;

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

        //ExStart:RevisionSpecifiedCriteria
        //GistId:470c0da51e4317baae82ad9495747fed
        //ExFor:RevisionCollection.Accept(IRevisionCriteria)
        //ExFor:RevisionCollection.Reject(IRevisionCriteria)
        //ExFor:IRevisionCriteria
        //ExFor:IRevisionCriteria.IsMatch(Revision)
        //ExSummary:Shows how to accept or reject revision based on criteria.
        [Test] //ExSkip
        public void RevisionSpecifiedCriteria()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("This does not count as a revision. ");

            // To register our edits as revisions, we need to declare an author, and then start tracking them.
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            builder.Write("This is insertion revision #1. ");
            doc.StopTrackRevisions();

            doc.StartTrackRevisions("Jane Doe", DateTime.Now);
            builder.Write("This is insertion revision #2. ");
            // Remove a run "This does not count as a revision.".
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            doc.StopTrackRevisions();

            Assert.AreEqual(3, doc.Revisions.Count);
            // We have two revisions from different authors, so we need to accept only one.
            doc.Revisions.Accept(new RevisionCriteria("John Doe", RevisionType.Insertion));
            Assert.AreEqual(2, doc.Revisions.Count);
            // Reject revision with different author name and revision type.
            doc.Revisions.Reject(new RevisionCriteria("Jane Doe", RevisionType.Deletion));
            Assert.AreEqual(1, doc.Revisions.Count);

            doc.Save(ArtifactsDir + "Revision.RevisionSpecifiedCriteria.docx");
        }

        /// <summary>
        /// Control when certain revision should be accepted/rejected.
        /// </summary>
        public class RevisionCriteria : IRevisionCriteria
        {
            private readonly string AuthorName;
            private readonly RevisionType RevisionType;

            public RevisionCriteria(string authorName, RevisionType revisionType)
            {
                AuthorName = authorName;
                RevisionType = revisionType;
            }

            public bool IsMatch(Revision revision)
            {
                return revision.Author == AuthorName && revision.RevisionType == RevisionType;
            }
        }
        //ExEnd:RevisionSpecifiedCriteria

        [Test]
        public void TrackRevisions()
        {
            //ExStart
            //ExFor:Document.StartTrackRevisions(String)
            //ExFor:Document.StartTrackRevisions(String, DateTime)
            //ExFor:Document.StopTrackRevisions
            //ExSummary:Shows how to track revisions while editing a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Editing a document usually does not count as a revision until we begin tracking them.
            builder.Write("Hello world! ");

            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.False(doc.FirstSection.Body.Paragraphs[0].Runs[0].IsInsertRevision);

            doc.StartTrackRevisions("John Doe");

            builder.Write("Hello again! ");

            Assert.AreEqual(1, doc.Revisions.Count);
            Assert.True(doc.FirstSection.Body.Paragraphs[0].Runs[1].IsInsertRevision);
            Assert.AreEqual("John Doe", doc.Revisions[0].Author);
            Assert.IsTrue((DateTime.Now - doc.Revisions[0].DateTime).Milliseconds <= 10);

            // Stop tracking revisions to not count any future edits as revisions.
            doc.StopTrackRevisions();
            builder.Write("Hello again! ");

            Assert.AreEqual(1, doc.Revisions.Count);
            Assert.False(doc.FirstSection.Body.Paragraphs[0].Runs[2].IsInsertRevision);

            // Creating revisions gives them a date and time of the operation.
            // We can disable this by passing DateTime.MinValue when we start tracking revisions.
            doc.StartTrackRevisions("John Doe", DateTime.MinValue);
            builder.Write("Hello again! ");

            Assert.AreEqual(2, doc.Revisions.Count);
            Assert.AreEqual("John Doe", doc.Revisions[1].Author);
            Assert.AreEqual(DateTime.MinValue, doc.Revisions[1].DateTime);

            // We can accept/reject these revisions programmatically
            // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
            // In Microsoft Word, we can process them manually via "Review" -> "Changes".
            doc.Save(ArtifactsDir + "Revision.StartTrackRevisions.docx");
            //ExEnd
        }

        [Test]
        public void AcceptAllRevisions()
        {
            //ExStart
            //ExFor:Document.AcceptAllRevisions
            //ExSummary:Shows how to accept all tracking changes in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Edit the document while tracking changes to create a few revisions.
            doc.StartTrackRevisions("John Doe");
            builder.Write("Hello world! ");
            builder.Write("Hello again! ");
            builder.Write("This is another revision.");
            doc.StopTrackRevisions();

            Assert.AreEqual(3, doc.Revisions.Count);

            // We can iterate through every revision and accept/reject it as a part of our document.
            // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
            doc.AcceptAllRevisions();

            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.AreEqual("Hello world! Hello again! This is another revision.", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void GetRevisedPropertiesOfList()
        {
            //ExStart
            //ExFor:RevisionsView
            //ExFor:Document.RevisionsView
            //ExSummary:Shows how to switch between the revised and the original view of a document.
            Document doc = new Document(MyDir + "Revisions at list levels.docx");
            doc.UpdateListLabels();

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            Assert.AreEqual("1.", paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual("a.", paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual(string.Empty, paragraphs[2].ListLabel.LabelString);

            // View the document object as if all the revisions are accepted. Currently supports list labels.
            doc.RevisionsView = RevisionsView.Final;

            Assert.AreEqual(string.Empty, paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual("1.", paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual("a.", paragraphs[2].ListLabel.LabelString);
            //ExEnd

            doc.RevisionsView = RevisionsView.Original;
            doc.AcceptAllRevisions();

            Assert.AreEqual("a.", paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual(string.Empty, paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual("b.", paragraphs[2].ListLabel.LabelString);
        }

        [Test]
        public void Compare()
        {
            //ExStart
            //ExFor:Document.Compare(Document, String, DateTime)
            //ExFor:RevisionCollection.AcceptAll
            //ExSummary:Shows how to compare documents.
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);
            builder.Writeln("This is the original document.");

            Document docEdited = new Document();
            builder = new DocumentBuilder(docEdited);
            builder.Writeln("This is the edited document.");

            // Comparing documents with revisions will throw an exception.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
                docOriginal.Compare(docEdited, "authorName", DateTime.Now);

            // After the comparison, the original document will gain a new revision
            // for every element that is different in the edited document.
            Assert.AreEqual(2, docOriginal.Revisions.Count); //ExSkip
            foreach (Revision r in docOriginal.Revisions)
            {
                Console.WriteLine($"Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
                Console.WriteLine($"\tChanged text: \"{r.ParentNode.GetText()}\"");
            }

            // Accepting these revisions will transform the original document into the edited document.
            docOriginal.Revisions.AcceptAll();

            Assert.AreEqual(docOriginal.GetText(), docEdited.GetText());
            //ExEnd

            docOriginal = DocumentHelper.SaveOpen(docOriginal);
            Assert.AreEqual(0, docOriginal.Revisions.Count);
        }

        [Test]
        public void CompareDocumentWithRevisions()
        {
            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);
            builder.Writeln("Hello world! This text is not a revision.");

            Document docWithRevision = new Document();
            builder = new DocumentBuilder(docWithRevision);

            docWithRevision.StartTrackRevisions("John Doe");
            builder.Writeln("This is a revision.");

            Assert.Throws<InvalidOperationException>(() => docWithRevision.Compare(doc1, "John Doe", DateTime.Now));
        }

        [Test]
        public void CompareOptions()
        {
            //ExStart
            //ExFor:CompareOptions
            //ExFor:CompareOptions.CompareMoves
            //ExFor:CompareOptions.IgnoreFormatting
            //ExFor:CompareOptions.IgnoreCaseChanges
            //ExFor:CompareOptions.IgnoreComments
            //ExFor:CompareOptions.IgnoreTables
            //ExFor:CompareOptions.IgnoreFields
            //ExFor:CompareOptions.IgnoreFootnotes
            //ExFor:CompareOptions.IgnoreTextboxes
            //ExFor:CompareOptions.IgnoreHeadersAndFooters
            //ExFor:CompareOptions.Target
            //ExFor:ComparisonTargetType
            //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
            //ExSummary:Shows how to filter specific types of document elements when making a comparison.
            // Create the original document and populate it with various kinds of elements.
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);

            // Paragraph text referenced with an endnote:
            builder.Writeln("Hello world! This is the first paragraph.");
            builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

            // Table:
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original cell 1 text");
            builder.InsertCell();
            builder.Write("Original cell 2 text");
            builder.EndTable();

            // Textbox:
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 20);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Original textbox contents");

            // DATE field:
            builder.MoveTo(docOriginal.FirstSection.Body.AppendParagraph(""));
            builder.InsertField(" DATE ");

            // Comment:
            Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", DateTime.Now);
            newComment.SetText("Original comment.");
            builder.CurrentParagraph.AppendChild(newComment);

            // Header:
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Original header contents.");

            // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
            Document docEdited = (Document)docOriginal.Clone(true);
            Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;

            firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
            firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];
            ((Footnote)docEdited.GetChild(NodeType.Footnote, 0, true)).FirstParagraph.Runs[1].Text = "Edited endnote text.";
            ((Table)docEdited.GetChild(NodeType.Table, 0, true)).FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";
            ((Shape)docEdited.GetChild(NodeType.Shape, 0, true)).FirstParagraph.Runs[0].Text = "Edited textbox contents";
            ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true;
            ((Comment)docEdited.GetChild(NodeType.Comment, 0, true)).FirstParagraph.Runs[0].Text = "Edited comment.";
            docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].FirstParagraph.Runs[0].Text =
                "Edited header contents.";

            // Comparing documents creates a revision for every edit in the edited document.
            // A CompareOptions object has a series of flags that can suppress revisions
            // on each respective type of element, effectively ignoring their change.
            CompareOptions compareOptions = new CompareOptions
            {
                CompareMoves = false,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };

            docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);
            docOriginal.Save(ArtifactsDir + "Revision.CompareOptions.docx");
            //ExEnd

            docOriginal = new Document(ArtifactsDir + "Revision.CompareOptions.docx");

            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "OriginalEdited endnote text.", (Footnote)docOriginal.GetChild(NodeType.Footnote, 0, true));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IgnoreDmlUniqueId(bool isIgnoreDmlUniqueId)
        {
            //ExStart
            //ExFor:CompareOptions.AdvancedOptions
            //ExFor:AdvancedCompareOptions.IgnoreDmlUniqueId
            //ExFor:CompareOptions.IgnoreDmlUniqueId
            //ExSummary:Shows how to compare documents ignoring DML unique ID.
            Document docA = new Document(MyDir + "DML unique ID original.docx");
            Document docB = new Document(MyDir + "DML unique ID compare.docx");

            // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
            // If we are ignoring DML's unique ID, and revisions count were 0.
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.AdvancedOptions.IgnoreDmlUniqueId = isIgnoreDmlUniqueId;

            docA.Compare(docB, "Aspose.Words", DateTime.Now, compareOptions);

            Assert.AreEqual(isIgnoreDmlUniqueId ? 0 : 2, docA.Revisions.Count);
            //ExEnd
        }

        [Test]
        public void LayoutOptionsRevisions()
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:LayoutOptions.RevisionOptions
            //ExFor:RevisionColor
            //ExFor:RevisionOptions
            //ExFor:RevisionOptions.InsertedTextColor
            //ExFor:RevisionOptions.ShowRevisionBars
            //ExFor:RevisionOptions.RevisionBarsPosition
            //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a revision, then change the color of all revisions to green.
            builder.Writeln("This is not a revision.");
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            Assert.AreEqual(RevisionColor.ByAuthor, doc.LayoutOptions.RevisionOptions.InsertedTextColor); //ExSkip
            Assert.True(doc.LayoutOptions.RevisionOptions.ShowRevisionBars); //ExSkip
            builder.Writeln("This is a revision.");
            doc.StopTrackRevisions();
            builder.Writeln("This is not a revision.");

            // Remove the bar that appears to the left of every revised line.
            doc.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;
            doc.LayoutOptions.RevisionOptions.ShowRevisionBars = false;
            doc.LayoutOptions.RevisionOptions.RevisionBarsPosition = HorizontalAlignment.Right;

            doc.Save(ArtifactsDir + "Revision.LayoutOptionsRevisions.pdf");
            //ExEnd
        }

        [TestCase(Granularity.CharLevel)]
        [TestCase(Granularity.WordLevel)]
        public void GranularityCompareOption(Granularity granularity)
        {
            //ExStart
            //ExFor:CompareOptions.Granularity
            //ExFor:Granularity
            //ExSummary:Shows to specify a granularity while comparing documents.
            Document docA = new Document();
            DocumentBuilder builderA = new DocumentBuilder(docA);
            builderA.Writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

            Document docB = new Document();
            DocumentBuilder builderB = new DocumentBuilder(docB);
            builderB.Writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

            // Specify whether changes are tracking
            // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.Granularity = granularity;

            docA.Compare(docB, "author", DateTime.Now, compareOptions);

            // The first document's collection of revision groups contains all the differences between documents.
            RevisionGroupCollection groups = docA.Revisions.Groups;
            Assert.AreEqual(5, groups.Count);
            //ExEnd

            if (granularity == Granularity.CharLevel)
            {
                Assert.AreEqual(RevisionType.Deletion, groups[0].RevisionType);
                Assert.AreEqual("Alpha ", groups[0].Text);

                Assert.AreEqual(RevisionType.Deletion, groups[1].RevisionType);
                Assert.AreEqual(",", groups[1].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[2].RevisionType);
                Assert.AreEqual("s", groups[2].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[3].RevisionType);
                Assert.AreEqual("- \"", groups[3].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[4].RevisionType);
                Assert.AreEqual("\"", groups[4].Text);
            }
            else
            {
                Assert.AreEqual(RevisionType.Deletion, groups[0].RevisionType);
                Assert.AreEqual("Alpha Lorem", groups[0].Text);

                Assert.AreEqual(RevisionType.Deletion, groups[1].RevisionType);
                Assert.AreEqual(",", groups[1].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[2].RevisionType);
                Assert.AreEqual("Lorems", groups[2].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[3].RevisionType);
                Assert.AreEqual("- \"", groups[3].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[4].RevisionType);
                Assert.AreEqual("\"", groups[4].Text);
            }
        }

        [Test]
        public void IgnoreStoreItemId()
        {
            //ExStart:IgnoreStoreItemId
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:AdvancedCompareOptions
            //ExFor:AdvancedCompareOptions.IgnoreStoreItemId
            //ExSummary:Shows how to compare SDT with same content but different store item id.
            Document docA = new Document(MyDir + "Document with SDT 1.docx");
            Document docB = new Document(MyDir + "Document with SDT 2.docx");

            // Configure options to compare SDT with same content but different store item id.
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.AdvancedOptions.IgnoreStoreItemId = false;

            docA.Compare(docB, "user", DateTime.Now, compareOptions);
            Assert.AreEqual(8, docA.Revisions.Count);

            compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

            docA.Revisions.RejectAll();
            docA.Compare(docB, "user", DateTime.Now, compareOptions);
            Assert.AreEqual(0, docA.Revisions.Count);
            //ExEnd:IgnoreStoreItemId
        }

        [Test]
        public void RevisionCellColor()
        {
            //ExStart:RevisionCellColor
            //GistId:366eb64fd56dec3c2eaa40410e594182
            //ExFor:RevisionOptions.InsertCellColor
            //ExFor:RevisionOptions.DeleteCellColor
            //ExSummary:Shows how to work with insert/delete cell revision color.
            Document doc = new Document(MyDir + "Cell revisions.docx");

            doc.LayoutOptions.RevisionOptions.InsertCellColor = RevisionColor.LightBlue;
            doc.LayoutOptions.RevisionOptions.DeleteCellColor = RevisionColor.DarkRed;

            doc.Save(ArtifactsDir + "Revision.RevisionCellColor.pdf");
            //ExEnd:RevisionCellColor
        }
    }
}
