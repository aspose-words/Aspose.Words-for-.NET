// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRange : ApiExampleBase
    {
        [Test]
        public void ReplaceSimple()
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExFor:FindReplaceOptions
            //ExFor:FindReplaceOptions.MatchCase
            //ExFor:FindReplaceOptions.FindWholeWordsOnly
            //ExSummary:Simple find and replace operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");

            // Check the document contains what we are about to test
            Console.WriteLine(doc.FirstSection.Body.Paragraphs[0].GetText());

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = false;

            doc.Range.Replace("_CustomerName_", "James Bond", options);

            doc.Save(ArtifactsDir + "Range.ReplaceSimple.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Range.ReplaceSimple.docx");

            Assert.AreEqual("Hello James Bond,", doc.GetText().Trim());
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreDeleted(bool isIgnoreDeleted)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreDeleted
            //ExSummary:Shows how to ignore text inside delete revisions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            // Insert non-revised text
            builder.Writeln("Deleted");
            builder.Write("Text");
 
            // Remove first paragraph with tracking revisions
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            doc.FirstSection.Body.FirstParagraph.Remove();
            doc.StopTrackRevisions();
 
            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();
 
            // Replace 'e' in document while ignoring/not ignoring deleted text
            options.IgnoreDeleted = isIgnoreDeleted;
            doc.Range.Replace(regex, "*", options);

            Assert.AreEqual(doc.GetText().Trim(), isIgnoreDeleted ? "Deleted\rT*xt" : "D*l*t*d\rT*xt");
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreInserted(bool isIgnoreInserted)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreInserted
            //ExSummary:Shows how to ignore text inside insert revisions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            // Insert text with tracking revisions
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            builder.Writeln("Inserted");
            doc.StopTrackRevisions();
 
            // Insert non-revised text
            builder.Write("Text");
 
            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();
 
            // Replace 'e' in document while ignoring/not ignoring inserted text
            options.IgnoreInserted = isIgnoreInserted;
            doc.Range.Replace(regex, "*", options);

            Assert.AreEqual(doc.GetText().Trim(), isIgnoreInserted ? "Inserted\rT*xt" : "Ins*rt*d\rT*xt");
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreFields(bool isIgnoreFields)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreFields
            //ExSummary:Shows how to ignore text inside fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            // Insert field with text inside
            builder.InsertField("INCLUDETEXT", "Text in field");
 
            Regex regex = new Regex("e");
            FindReplaceOptions options = new FindReplaceOptions();
            // Replace 'e' in document ignoring/not ignoring text inside field
            options.IgnoreFields = isIgnoreFields;
            
            doc.Range.Replace(regex, "*", options);

            Assert.AreEqual(doc.GetText(),
                isIgnoreFields
                    ? "\u0013INCLUDETEXT\u0014Text in field\u0015\f"
                    : "\u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f");
            //ExEnd
        }

        [Test]
        public void UpdateFieldsInRange()
        {
            //ExStart
            //ExFor:Range.UpdateFields
            //ExSummary:Shows how to update document fields in the body of the first section only.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a field that will display the value in the document's body text
            FieldDocProperty field = (FieldDocProperty)builder.InsertField(" DOCPROPERTY Category");

            // Set the value of the property that should be displayed by the field
            doc.BuiltInDocumentProperties.Category = "MyCategory";

            // Some field types need to be explicitly updated before they can display their expected values
            Assert.AreEqual(string.Empty, field.Result);

            // Update all the fields in the first section of the document, which includes the field we just inserted
            doc.FirstSection.Range.UpdateFields();

            Assert.AreEqual("MyCategory", field.Result);
            //ExEnd
        }

        [Test]
        public void ReplaceWithString()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This one is sad.");
            builder.Writeln("That one is mad.");

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = true;

            doc.Range.Replace("sad", "bad", options);

            doc.Save(ArtifactsDir + "Range.ReplaceWithString.docx");
        }

        [Test]
        public void ReplaceWithRegex()
        {
            //ExStart
            //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("sad mad bad");

            Assert.AreEqual("sad mad bad", doc.GetText().Trim());

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);

            Assert.AreEqual("bad bad bad", doc.GetText().Trim());
            //ExEnd
        }

        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:ReplacingArgs.Replacement
        //ExFor:IReplacingCallback
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExSummary:Replaces text specified with regular expression with HTML.
        [Test] //ExSkip
        public void ReplaceWithInsertHtml()
        {
            // Open the document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello <CustomerName>,");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithHtmlEvaluator();

            doc.Range.Replace(new Regex(@" <CustomerName>,"), string.Empty, options);

            // Save the modified document
            doc.Save(ArtifactsDir + "Range.ReplaceWithInsertHtml.docx");
            Assert.AreEqual("James Bond, Hello\r\x000c",
                new Document(ArtifactsDir + "Range.ReplaceWithInsertHtml.docx").GetText()); //ExSkip
        }

        private class ReplaceWithHtmlEvaluator : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // Replace '<CustomerName>' text with a red bold name
                builder.InsertHtml("<b><font color='red'>James Bond, </font></b>");
                args.Replacement = "";

                return ReplaceAction.Replace;
            }
        }
        //ExEnd

        //ExStart
        //ExFor:FindReplaceOptions.ApplyFont
        //ExFor:FindReplaceOptions.Direction
        //ExFor:FindReplaceOptions.ReplacingCallback
        //ExFor:ReplacingArgs.GroupIndex
        //ExFor:ReplacingArgs.GroupName
        //ExFor:ReplacingArgs.Match
        //ExFor:ReplacingArgs.MatchOffset
        //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
        [Test] //ExSkip
        public void ReplaceNumbersAsHex()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Numbers that will be converted to hexadecimal and highlighted:\n" +
                            "123, 456, 789 and 17379.");

            FindReplaceOptions options = new FindReplaceOptions();
            // Highlight newly inserted content with a color
            options.ApplyFont.HighlightColor = Color.LightGray;
            // Apply an IReplacingCallback to make the replacement to convert integers into hex equivalents
            // and also to count replacements in the order they take place
            options.ReplacingCallback = new NumberHexer();
            // By default, text is searched for replacements front to back, but we can change it to go the other way
            options.Direction = FindReplaceDirection.Backward;

            int count = doc.Range.Replace(new Regex("[0-9]+"), "", options);

            Assert.AreEqual(4, count);
            Assert.AreEqual("Numbers that will be converted to hexadecimal and highlighted:\r" +
                            "0x7B, 0x1C8, 0x315 and 0x43E3.", doc.GetText().Trim());
            Assert.AreEqual(4,
                doc.GetChildNodes(NodeType.Run, true).OfType<Run>()
                    .Count(r => r.Font.HighlightColor.ToArgb() == Color.LightGray.ToArgb()));
        }

        /// <summary>
        /// Replaces arabic numbers with hexadecimal equivalents and appends the number of each replacement.
        /// </summary>
        private class NumberHexer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                mCurrentReplacementNumber++;
                
                // Parse numbers
                int number = Convert.ToInt32(args.Match.Value);

                // And write it as HEX
                args.Replacement = $"0x{number:X}";

                Console.WriteLine($"Match #{mCurrentReplacementNumber}");
                Console.WriteLine($"\tOriginal value:\t{args.Match.Value}");
                Console.WriteLine($"\tReplacement:\t{args.Replacement}");
                Console.WriteLine($"\tOffset in parent {args.MatchNode.NodeType} node:\t{args.MatchOffset}");

                Console.WriteLine(string.IsNullOrEmpty(args.GroupName)
                    ? $"\tGroup index:\t{args.GroupIndex}"
                    : $"\tGroup name:\t{args.GroupName}");

                return ReplaceAction.Replace;
            }

            private int mCurrentReplacementNumber;
        }
        //ExEnd

        [Test]
        public void ApplyParagraphFormat()
        {
            //ExStart
            //ExFor:FindReplaceOptions.ApplyParagraphFormat
            //ExSummary:Shows how to affect the format of paragraphs with successful replacements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
            builder.Writeln("This one will not!");
            builder.Writeln("And this one will.");
            
            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;

            int count = doc.Range.Replace(".&p", "!&p", options);
            Assert.AreEqual(2, count);

            doc.Save(ArtifactsDir + "Range.ApplyParagraphFormat.docx");
            //ExEnd

            ParagraphCollection paragraphs = new Document(ArtifactsDir + "Range.ApplyParagraphFormat.docx").FirstSection.Body.Paragraphs;

            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[0].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[1].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[2].ParagraphFormat.Alignment);
        }

        [Test]
        public void DeleteSelection()
        {
            //ExStart
            //ExFor:Node.Range
            //ExFor:Range.Delete
            //ExSummary:Shows how to delete all characters of a range.
            // Insert two sections into a blank document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1. ");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Write("Section 2.");

            // Verify the whole text of the document
            Assert.AreEqual("Section 1. \fSection 2.", doc.GetText().Trim());

            // Delete the first section from the document
            doc.Sections[0].Range.Delete();

            // Check the first section was deleted by looking at the text of the whole document again
            Assert.AreEqual("Section 2.", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void RangesGetText()
        {
            //ExStart
            //ExFor:Range
            //ExFor:Range.Text
            //ExSummary:Shows how to get plain, unformatted text of a range.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            Assert.AreEqual("Hello world!", doc.Range.Text.Trim());
            //ExEnd
        }

        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:IReplacingCallback
        //ExFor:ReplaceAction
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:ReplacingArgs.MatchNode
        //ExFor:FindReplaceDirection
        //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
        [Test] //ExSkip
        public void InsertDocumentAtReplace()
        {
            Document mainDoc = new Document(MyDir + "Document insertion destination.docx");

            FindReplaceOptions options = new FindReplaceOptions();
            options.Direction = FindReplaceDirection.Backward;
            options.ReplacingCallback = new InsertDocumentAtReplaceHandler();

            mainDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"), "", options);
            mainDoc.Save(ArtifactsDir + "InsertDocument.InsertDocumentAtReplace.docx");

            TestInsertDocumentAtReplace(new Document(ArtifactsDir + "InsertDocument.InsertDocumentAtReplace.docx")); //ExSkip
        }

        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                Document subDoc = new Document(MyDir + "Document.docx");

                // Insert a document after the paragraph, containing the match text
                Paragraph para = (Paragraph)args.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                // Remove the paragraph with the match text
                para.Remove();

                return ReplaceAction.Skip;
            }
        }

        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// </summary>
        static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // Make sure that the node is either a paragraph or table
            if (insertionDestination.NodeType.Equals(NodeType.Paragraph) || insertionDestination.NodeType.Equals(NodeType.Table))
            {
                // We will be inserting into the parent of the destination paragraph
                CompositeNode dstStory = insertionDestination.ParentNode;

                // This object will be translating styles and lists during the import
                NodeImporter importer =
                    new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

                // Loop through all block level nodes in the body of the section
                foreach (Section srcSection in docToInsert.Sections.OfType<Section>())
                    foreach (Node srcNode in srcSection.Body)
                    {
                        // Let's skip the node if it is a last empty paragraph in a section
                        if (srcNode.NodeType.Equals(NodeType.Paragraph))
                        {
                            Paragraph para = (Paragraph)srcNode;
                            if (para.IsEndOfSection && !para.HasChildNodes)
                                continue;
                        }

                        // This creates a clone of the node, suitable for insertion into the destination document
                        Node newNode = importer.ImportNode(srcNode, true);

                        // Insert new node after the reference node
                        dstStory.InsertAfter(newNode, insertionDestination);
                        insertionDestination = newNode;
                    }
            }
            else
            {
                throw new ArgumentException("The destination node should be either a paragraph or table.");
            }
        }
        //ExEnd

        private void TestInsertDocumentAtReplace(Document doc)
        {
            Assert.AreEqual("1) At text that can be identified by regex:\rHello World!\r" +
                            "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                            "3) At a bookmark:", doc.FirstSection.Body.GetText().Trim());
        }
    }
}