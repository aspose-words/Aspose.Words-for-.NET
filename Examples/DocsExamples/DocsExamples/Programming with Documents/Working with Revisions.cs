using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithRevisions : DocsExamplesBase
    {
        [Test]
        public void AcceptRevisions()
        {
            //ExStart:AcceptAllRevisions
            Document doc = new Document();
            Body body = doc.FirstSection.Body;
            Paragraph para = body.FirstParagraph;

            // Add text to the first paragraph, then add two more paragraphs.
            para.AppendChild(new Run(doc, "Paragraph 1. "));
            body.AppendParagraph("Paragraph 2. ");
            body.AppendParagraph("Paragraph 3. ");

            // We have three paragraphs, none of which registered as any type of revision
            // If we add/remove any content in the document while tracking revisions,
            // they will be displayed as such in the document and can be accepted/rejected.
            doc.StartTrackRevisions("John Doe", DateTime.Now);

            // This paragraph is a revision and will have the according "IsInsertRevision" flag set.
            para = body.AppendParagraph("Paragraph 4. ");
            Assert.True(para.IsInsertRevision);

            // Get the document's paragraph collection and remove a paragraph.
            ParagraphCollection paragraphs = body.Paragraphs;
            Assert.AreEqual(4, paragraphs.Count);
            para = paragraphs[2];
            para.Remove();

            // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
            // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
            Assert.AreEqual(4, paragraphs.Count);
            Assert.True(para.IsDeleteRevision);

            // The delete revision paragraph is removed once we accept changes.
            doc.AcceptAllRevisions();
            Assert.AreEqual(3, paragraphs.Count);
            Assert.That(para, Is.Empty);

            // Stopping the tracking of revisions makes this text appear as normal text.
            // Revisions are not counted when the document is changed.
            doc.StopTrackRevisions();

            // Save the document.
            doc.Save(ArtifactsDir + "WorkingWithRevisions.AcceptRevisions.docx");
            //ExEnd:AcceptAllRevisions
        }

        [Test]
        public void GetRevisionTypes()
        {
            //ExStart:GetRevisionTypes
            Document doc = new Document(MyDir + "Revisions.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            //ExEnd:GetRevisionTypes
        }

        [Test]
        public void GetRevisionGroups()
        {
            //ExStart:GetRevisionGroups
            Document doc = new Document(MyDir + "Revisions.docx");

            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine("{0}, {1}:", group.Author, group.RevisionType);
                Console.WriteLine(group.Text);
            }
            //ExEnd:GetRevisionGroups
        }

        [Test]
        public void RemoveCommentsInPdf()
        {
            //ExStart:RemoveCommentsInPDF
            Document doc = new Document(MyDir + "Revisions.docx");

            // Do not render the comments in PDF.
            doc.LayoutOptions.ShowComments = false;

            doc.Save(ArtifactsDir + "WorkingWithRevisions.RemoveCommentsInPdf.pdf");
            //ExEnd:RemoveCommentsInPDF
        }

        [Test]
        public void ShowRevisionsInBalloons()
        {
            //ExStart:ShowRevisionsInBalloons
            //ExStart:SetMeasurementUnit
            //ExStart:SetRevisionBarsPosition
            Document doc = new Document(MyDir + "Revisions.docx");

            // Renders insert and delete revisions inline, format revisions in balloons.
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.Format;
            doc.LayoutOptions.RevisionOptions.MeasurementUnit = MeasurementUnits.Inches;
            
            // Renders revision bars on the right side of a page.
            doc.LayoutOptions.RevisionOptions.RevisionBarsPosition = HorizontalAlignment.Right;

            // Renders insert revisions inline, delete and format revisions in balloons.
            //doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;

            doc.Save(ArtifactsDir + "WorkingWithRevisions.ShowRevisionsInBalloons.pdf");
            //ExEnd:SetRevisionBarsPosition
            //ExEnd:SetMeasurementUnit
            //ExEnd:ShowRevisionsInBalloons
        }

        [Test]
        public void GetRevisionGroupDetails()
        {
            //ExStart:GetRevisionGroupDetails
            Document doc = new Document(MyDir + "Revisions.docx");

            foreach (Revision revision in doc.Revisions)
            {
                string groupText = revision.Group != null
                    ? "Revision group text: " + revision.Group.Text
                    : "Revision has no group";

                Console.WriteLine("Type: " + revision.RevisionType);
                Console.WriteLine("Author: " + revision.Author);
                Console.WriteLine("Date: " + revision.DateTime);
                Console.WriteLine("Revision text: " + revision.ParentNode.ToString(SaveFormat.Text));
                Console.WriteLine(groupText);
            }
            //ExEnd:GetRevisionGroupDetails
        }

        [Test]
        public void AccessRevisedVersion()
        {
            //ExStart:AccessRevisedVersion
            Document doc = new Document(MyDir + "Revisions.docx");
            doc.UpdateListLabels();

            // Switch to the revised version of the document.
            doc.RevisionsView = RevisionsView.Final;

            foreach (Revision revision in doc.Revisions)
            {
                if (revision.ParentNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph paragraph = (Paragraph) revision.ParentNode;
                    if (paragraph.IsListItem)
                    {
                        Console.WriteLine(paragraph.ListLabel.LabelString);
                        Console.WriteLine(paragraph.ListFormat.ListLevel);
                    }
                }
            }
            //ExEnd:AccessRevisedVersion
        }

        [Test]
        public void MoveNodeInTrackedDocument()
        {
            //ExStart:MoveNodeInTrackedDocument
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2");
            builder.Writeln("Paragraph 3");
            builder.Writeln("Paragraph 4");
            builder.Writeln("Paragraph 5");
            builder.Writeln("Paragraph 6");
            Body body = doc.FirstSection.Body;
            Console.WriteLine("Paragraph count: {0}", body.Paragraphs.Count);

            // Start tracking revisions.
            doc.StartTrackRevisions("Author", new DateTime(2020, 12, 23, 14, 0, 0));

            // Generate revisions when moving a node from one location to another.
            Node node = body.Paragraphs[3];
            Node endNode = body.Paragraphs[5].NextSibling;
            Node referenceNode = body.Paragraphs[0];
            while (node != endNode)
            {
                Node nextNode = node.NextSibling;
                body.InsertBefore(node, referenceNode);
                node = nextNode;
            }

            // Stop the process of tracking revisions.
            doc.StopTrackRevisions();

            // There are 3 additional paragraphs in the move-from range.
            Console.WriteLine("Paragraph count: {0}", body.Paragraphs.Count);
            doc.Save(ArtifactsDir + "WorkingWithRevisions.MoveNodeInTrackedDocument.docx");
            //ExEnd:MoveNodeInTrackedDocument
        }

        [Test]
        public void ShapeRevision()
        {
            //ExStart:ShapeRevision
            Document doc = new Document();

            // Insert an inline shape without tracking revisions.
            Assert.False(doc.TrackRevisions);
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Start tracking revisions and then insert another shape.
            doc.StartTrackRevisions("John Doe");
            shape = new Shape(doc, ShapeType.Sun);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Get the document's shape collection which includes just the two shapes we added.
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // Remove the first shape.
            shapes[0].Remove();

            // Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
            Assert.AreEqual(ShapeType.Cube, shapes[0].ShapeType);
            Assert.True(shapes[0].IsDeleteRevision);

            // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
            Assert.AreEqual(ShapeType.Sun, shapes[1].ShapeType);
            Assert.True(shapes[1].IsInsertRevision);

            // The document has one shape that was moved, but shape move revisions will have two instances of that shape.
            // One will be the shape at its arrival destination and the other will be the shape at its original location.
            doc = new Document(MyDir + "Revision shape.docx");
            
            shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(4, shapes.Count);

            // This is the move to revision, also the shape at its arrival destination.
            Assert.False(shapes[0].IsMoveFromRevision);
            Assert.True(shapes[0].IsMoveToRevision);

            // This is the move from revision, which is the shape at its original location.
            Assert.True(shapes[1].IsMoveFromRevision);
            Assert.False(shapes[1].IsMoveToRevision);
            //ExEnd:ShapeRevision
        }
    }
}