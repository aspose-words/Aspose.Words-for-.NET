
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Layout;
using System.Text.RegularExpressions;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRevisions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            AcceptRevisions(dataDir);
            GetRevisionTypes(dataDir);
            GetRevisionGroups(dataDir);
            SetShowCommentsinPDF(dataDir);
            GetRevisionGroupDetails(dataDir);
            AccessRevisedVersion(dataDir);
            IgnoreTextInsideDeleteRevisions(dataDir);
            IgnoreTextInsideInsertRevisions(dataDir);
        }

        private static void AcceptRevisions(string dataDir)
        {
            // ExStart:AcceptAllRevisions
            Document doc = new Document(dataDir + "Document.doc");

            // Start tracking and make some revisions.
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document.
            doc.AcceptAllRevisions();

            dataDir = dataDir + "Document.AcceptedRevisions_out.doc";
            doc.Save(dataDir);
            // ExEnd:AcceptAllRevisions
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }

        private static void GetRevisionTypes(string dataDir)
        {
            // ExStart:GetRevisionTypes
            Document doc = new Document(dataDir + "Revisions.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            // ExEnd:GetRevisionTypes
        }

        private static void GetRevisionGroups(string dataDir)
        {
            // ExStart:GetRevisionGroups
            Document doc = new Document(dataDir + "Revisions.docx");

            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine("{0}, {1}:", group.Author, group.RevisionType);
                Console.WriteLine(group.Text);
            }
            // ExEnd:GetRevisionGroups
        }

        private static void SetShowCommentsinPDF(string dataDir)
        {
            // ExStart:SetShowCommentsinPDF
            Document doc = new Document(dataDir + "Revisions.docx");

            //Do not render the comments in PDF
            doc.LayoutOptions.ShowComments = false;
            doc.Save(dataDir + "RemoveCommentsinPDF_out.pdf");
            // ExEnd:SetShowCommentsinPDF
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        private static void SetShowInBalloons(string dataDir)
        {
            // ExStart:SetShowInBalloons
            Document doc = new Document(dataDir + "Revisions.docx");

            // Renders insert and delete revisions inline, format revisions in balloons.
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.Format;

            // Renders insert revisions inline, delete and format revisions in balloons.
            //doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.FormatAndDelete;

            doc.Save(dataDir + "SetShowInBalloons_out.pdf");
            // ExEnd:SetShowInBalloons
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        private static void GetRevisionGroupDetails(string dataDir)
        {
            // ExStart:GetRevisionGroupDetails
            Document doc = new Document(dataDir + "TestFormatDescription.docx");

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
            // ExEnd:GetRevisionGroupDetails
        }

        private static void AccessRevisedVersion(string dataDir)
        {
            // ExStart:AccessRevisedVersion
            Document doc = new Document(dataDir + "Test.docx");
            doc.UpdateListLabels();

            // Switch to the revised version of the document.
            doc.RevisionsView = RevisionsView.Final;

            foreach (Revision revision in doc.Revisions)
            {
                if (revision.ParentNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph paragraph = (Paragraph)revision.ParentNode;
                    if (paragraph.IsListItem)
                    {
                        // Print revised version of LabelString and ListLevel.
                        Console.WriteLine(paragraph.ListLabel.LabelString);
                        Console.WriteLine(paragraph.ListFormat.ListLevel);
                    }
                }
            }
            // ExEnd:AccessRevisedVersion
        }

        private static void IgnoreTextInsideDeleteRevisions(string dataDir)
        {
            // ExStart:IgnoreTextInsideDeleteRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert non-revised text.
            builder.Writeln("Deleted");
            builder.Write("Text");

            // Remove first paragraph with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            doc.FirstSection.Body.FirstParagraph.Remove();
            doc.StopTrackRevisions();

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring deleted text.
            options.IgnoreDeleted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Deleted\rT*xt\f

            // Replace 'e' in document NOT ignoring deleted text.
            options.IgnoreDeleted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: D*l*t*d\rT*xt\f
            // ExEnd:IgnoreTextInsideDeleteRevisions
        }

        private static void IgnoreTextInsideInsertRevisions(string dataDir)
        {
            // ExStart:IgnoreTextInsideInsertRevisions
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            builder.Writeln("Inserted");
            doc.StopTrackRevisions();

            // Insert non-revised text.
            builder.Write("Text");

            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();

            // Replace 'e' in document ignoring inserted text.
            options.IgnoreInserted = true;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Inserted\rT*xt\f

            // Replace 'e' in document NOT ignoring inserted text.
            options.IgnoreInserted = false;
            doc.Range.Replace(regex, "*", options);
            Console.WriteLine(doc.GetText()); // The output is: Ins*rt*d\rT*xt\f
            // ExEnd:IgnoreTextInsideInsertRevisions
        }
    }
}
