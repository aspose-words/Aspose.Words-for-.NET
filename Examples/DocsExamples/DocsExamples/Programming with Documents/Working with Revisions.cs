using System;
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
            Document doc = new Document(MyDir + "Document.docx");

            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document.
            doc.AcceptAllRevisions();
            
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
    }
}