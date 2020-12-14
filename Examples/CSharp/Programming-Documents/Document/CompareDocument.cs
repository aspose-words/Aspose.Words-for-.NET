
using System.IO;
using Aspose.Words;
using System;
using System.Linq;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CompareDocument
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            NormalComparison(dataDir);
            CompareForEqual(dataDir);
            CompareDocumentWithCompareOptions(dataDir);
            CompareDocumentWithComparisonTarget(dataDir);
            SpecifyComparisonGranularity(dataDir);
            ApplyCompareTwoDocuments(dataDir);
            SetAdvancedComparingProperties(dataDir);
        }             

        private static void NormalComparison(string dataDir)
        {
            // ExStart:NormalComparison
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now); 
            // ExEnd:NormalComparison                     
        }
        private static void CompareForEqual(string dataDir)
        {
            // ExStart:CompareForEqual
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now);
            if (docA.Revisions.Count == 0)
                Console.WriteLine("Documents are equal");
            else
                Console.WriteLine("Documents are not equal");
            // ExEnd:CompareForEqual                     
        }

        private static void CompareDocumentWithCompareOptions(string dataDir)
        {
            // ExStart:CompareDocumentWithCompareOptions
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            options.IgnoreHeadersAndFooters = true;
            options.IgnoreCaseChanges = true;
            options.IgnoreTables = true;
            options.IgnoreFields = true;
            options.IgnoreComments = true;
            options.IgnoreTextboxes = true;
            options.IgnoreFootnotes = true;

            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now, options);
            if (docA.Revisions.Count == 0)
                Console.WriteLine("Documents are equal");
            else
                Console.WriteLine("Documents are not equal");
            // ExEnd:CompareDocumentWithCompareOptions                     
        }

        private static void CompareDocumentWithComparisonTarget(string dataDir)
        {
            // ExStart:CompareDocumentWithComparisonTarget
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box. 
            options.Target = ComparisonTargetType.New;

            docA.Compare(docB, "user", DateTime.Now, options);
            // ExEnd:CompareDocumentWithComparisonTarget      

            dataDir = dataDir + "TestFile_Out.doc";

            Console.WriteLine("\nDocuments have compared successfully.\nFile saved at " + dataDir);
        }

        public static void SpecifyComparisonGranularity(string dataDir)
        {
            // ExStart:SpecifyComparisonGranularity
            DocumentBuilder builderA = new DocumentBuilder(new Document());
            DocumentBuilder builderB = new DocumentBuilder(new Document());

            builderA.Writeln("This is A simple word");
            builderB.Writeln("This is B simple words");

            CompareOptions co = new CompareOptions();
            co.Granularity = Granularity.CharLevel;

            builderA.Document.Compare(builderB.Document, "author", DateTime.Now, co);
            // ExEnd:SpecifyComparisonGranularity
        }

        public static void ApplyCompareTwoDocuments(string dataDir)
        {
            //ExStart:ApplyCompareTwoDocuments
            // The source document doc1.
            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);
            builder.Writeln("This is the original document.");

            // The target document doc2.
            Document doc2 = new Document();
            builder = new DocumentBuilder(doc2);
            builder.Writeln("This is the edited document.");

            // If either document has a revision, an exception will be thrown.
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "authorName", DateTime.Now);

            // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed.
            Assert.AreEqual(2, doc1.Revisions.Count);

            foreach (Revision r in doc1.Revisions)
            {
                Console.WriteLine($"Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
                Console.WriteLine($"\tChanged text: \"{r.ParentNode.GetText()}\"");
            }

            // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
            doc1.Revisions.AcceptAll();

            // doc1, when saved, now resembles doc2.
            doc1.Save(dataDir + "Document.Compare.docx");
            doc1 = new Document(dataDir + "Document.Compare.docx");
            Assert.AreEqual(0, doc1.Revisions.Count);
            Assert.AreEqual(doc2.GetText().Trim(), doc1.GetText().Trim());
            //ExEnd:ApplyCompareTwoDocuments
        }

        public static void SetAdvancedComparingProperties(string dataDir)
        {
            //ExStart:SetAdvancedComparingProperties
            // Create the original document.
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);

            // Insert paragraph text with an endnote.
            builder.Writeln("Hello world! This is the first paragraph.");
            builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

            // Insert a table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original cell 1 text");
            builder.InsertCell();
            builder.Write("Original cell 2 text");
            builder.EndTable();

            // Insert a textbox.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 20);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Original textbox contents");

            // Insert a DATE field.
            builder.MoveTo(docOriginal.FirstSection.Body.AppendParagraph(""));
            builder.InsertField(" DATE ");

            // Insert a comment.
            Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", DateTime.Now);
            newComment.SetText("Original comment.");
            builder.CurrentParagraph.AppendChild(newComment);

            // Insert a header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Original header contents.");

            // Create a clone of our document, which we will edit and later compare to the original.
            Document docEdited = (Document)docOriginal.Clone(true);
            Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;

            // Change the formatting of the first paragraph, change casing of original characters and add text.
            firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
            firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];

            // Edit the footnote.
            Footnote footnote = (Footnote)docEdited.GetChild(NodeType.Footnote, 0, true);
            footnote.FirstParagraph.Runs[1].Text = "Edited endnote text.";

            // Edit the table.
            Table table = (Table)docEdited.GetChild(NodeType.Table, 0, true);
            table.FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";

            // Edit the textbox.
            textBox = (Shape)docEdited.GetChild(NodeType.Shape, 0, true);
            textBox.FirstParagraph.Runs[0].Text = "Edited textbox contents";

            // Edit the DATE field.
            FieldDate fieldDate = (FieldDate)docEdited.Range.Fields[0];
            fieldDate.UseLunarCalendar = true;

            // Edit the comment.
            Comment comment = (Comment)docEdited.GetChild(NodeType.Comment, 0, true);
            comment.FirstParagraph.Runs[0].Text = "Edited comment.";

            // Edit the header.
            docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].FirstParagraph.Runs[0].Text = "Edited header contents.";

            // Apply different comparing options.
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreFormatting = false;
            compareOptions.IgnoreCaseChanges = false;
            compareOptions.IgnoreComments = false;
            compareOptions.IgnoreTables = false;
            compareOptions.IgnoreFields = false;
            compareOptions.IgnoreFootnotes = false;
            compareOptions.IgnoreTextboxes = false;
            compareOptions.IgnoreHeadersAndFooters = false;
            compareOptions.Target = ComparisonTargetType.New;

            // compare both documents.
            docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);
            docOriginal.Save(dataDir + "Document.CompareOptions.docx");

            docOriginal = new Document(dataDir + "Document.CompareOptions.docx");

            // If you set compareOptions to ignore certain types of changes,
            // then revisions done on those types of nodes will not appear in the output document.
            // You can tell what kind of node a revision was done on by looking at the NodeType of the revision's parent nodes.
            Assert.AreNotEqual(compareOptions.IgnoreFormatting, docOriginal.Revisions.Any(rev => rev.RevisionType == RevisionType.FormatChange));
            Assert.AreNotEqual(compareOptions.IgnoreCaseChanges, docOriginal.Revisions.Any(s => s.ParentNode.GetText().Contains("hello")));
            Assert.AreNotEqual(compareOptions.IgnoreComments, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Comment)));
            Assert.AreNotEqual(compareOptions.IgnoreTables, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Table)));
            Assert.AreNotEqual(compareOptions.IgnoreFields, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.FieldStart)));
            Assert.AreNotEqual(compareOptions.IgnoreFootnotes, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Footnote)));
            Assert.AreNotEqual(compareOptions.IgnoreTextboxes, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Shape)));
            Assert.AreNotEqual(compareOptions.IgnoreHeadersAndFooters, docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.HeaderFooter)));
            //ExEnd:SetAdvancedComparingProperties
        }

        /// <summary>
        /// Returns true if the passed revision has a parent node with the type specified by parentType.
        /// </summary>
        private static bool HasParentOfType(Revision revision, NodeType parentType)
        {
            Node n = revision.ParentNode;
            while (n.ParentNode != null)
            {
                if (n.NodeType == parentType) return true;
                n = n.ParentNode;
            }

            return false;
        }
    }
}
