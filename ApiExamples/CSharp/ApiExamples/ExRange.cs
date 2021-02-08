// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRange : ApiExampleBase
    {
        [Test]
        public void Replace()
        {
            //ExStart
            //ExFor:Range.Replace(String, String)
            //ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Greetings, _FullName_!");

            // Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
            int replacementCount = doc.Range.Replace("_FullName_", "John Doe");

            Assert.AreEqual(1, replacementCount);
            Assert.AreEqual("Greetings, John Doe!", doc.GetText().Trim());
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ReplaceMatchCase(bool matchCase)
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExFor:FindReplaceOptions
            //ExFor:FindReplaceOptions.MatchCase
            //ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Ruby bought a ruby necklace.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
            // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
            options.MatchCase = matchCase;

            doc.Range.Replace("Ruby", "Jade", options);

            Assert.AreEqual(matchCase ? "Jade bought a ruby necklace." : "Jade bought a Jade necklace.",
                doc.GetText().Trim());
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ReplaceFindWholeWordsOnly(bool findWholeWordsOnly)
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExFor:FindReplaceOptions
            //ExFor:FindReplaceOptions.FindWholeWordsOnly
            //ExSummary:Shows how to toggle standalone word-only find-and-replace operations. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jackson will meet you in Jacksonville.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
            // Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
            options.FindWholeWordsOnly = findWholeWordsOnly;

            doc.Range.Replace("Jackson", "Louis", options);

            Assert.AreEqual(
                findWholeWordsOnly ? "Louis will meet you in Jacksonville." : "Louis will meet you in Louisville.",
                doc.GetText().Trim());
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreDeleted(bool ignoreTextInsideDeleteRevisions)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreDeleted
            //ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            builder.Writeln("Hello world!");
            builder.Writeln("Hello again!");
 
            // Start tracking revisions and remove the second paragraph, which will create a delete revision.
            // That paragraph will persist in the document until we accept the delete revision.
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            doc.FirstSection.Body.Paragraphs[1].Remove();
            doc.StopTrackRevisions();

            Assert.True(doc.FirstSection.Body.Paragraphs[1].IsDeleteRevision);

            // We can use a "FindReplaceOptions" object to modify the find and replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
            // operation to ignore paragraphs that are delete revisions.
            // Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
            // operation to also search for text inside delete revisions.
            options.IgnoreDeleted = ignoreTextInsideDeleteRevisions;
            
            doc.Range.Replace("Hello", "Greetings", options);

            Assert.AreEqual(
                ignoreTextInsideDeleteRevisions
                    ? "Greetings world!\rHello again!"
                    : "Greetings world!\rGreetings again!", doc.GetText().Trim());
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreInserted(bool ignoreTextInsideInsertRevisions)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreInserted
            //ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            // Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            builder.Writeln("Hello again!");
            doc.StopTrackRevisions();

            Assert.True(doc.FirstSection.Body.Paragraphs[1].IsInsertRevision);

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "IgnoreInserted" flag to "true" to get the find-and-replace
            // operation to ignore paragraphs that are insert revisions.
            // Set the "IgnoreInserted" flag to "false" to get the find-and-replace
            // operation to also search for text inside insert revisions.
            options.IgnoreInserted = ignoreTextInsideInsertRevisions;

            doc.Range.Replace("Hello", "Greetings", options);

            Assert.AreEqual(
                ignoreTextInsideInsertRevisions
                    ? "Greetings world!\rHello again!"
                    : "Greetings world!\rGreetings again!", doc.GetText().Trim());
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreFields(bool ignoreTextInsideFields)
        {
            //ExStart
            //ExFor:FindReplaceOptions.IgnoreFields
            //ExSummary:Shows how to ignore text inside fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.InsertField("QUOTE", "Hello again!");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "IgnoreFields" flag to "true" to get the find-and-replace
            // operation to ignore text inside fields.
            // Set the "IgnoreFields" flag to "false" to get the find-and-replace
            // operation to also search for text inside fields.
            options.IgnoreFields = ignoreTextInsideFields;

            doc.Range.Replace("Hello", "Greetings", options);

            Assert.AreEqual(
                ignoreTextInsideFields
                    ? "Greetings world!\r\u0013QUOTE\u0014Hello again!\u0015"
                    : "Greetings world!\r\u0013QUOTE\u0014Greetings again!\u0015", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void UpdateFieldsInRange()
        {
            //ExStart
            //ExFor:Range.UpdateFields
            //ExSummary:Shows how to update all the fields in a range.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(" DOCPROPERTY Category");
            builder.InsertBreak(BreakType.SectionBreakEvenPage);
            builder.InsertField(" DOCPROPERTY Category");

            // The above DOCPROPERTY fields will display the value of this built-in document property.
            doc.BuiltInDocumentProperties.Category = "MyCategory";

            // If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
            Assert.AreEqual(string.Empty, doc.Range.Fields[0].Result);
            Assert.AreEqual(string.Empty, doc.Range.Fields[1].Result);

            // Update all the fields that are in the range of the first section.
            doc.FirstSection.Range.UpdateFields();

            Assert.AreEqual("MyCategory", doc.Range.Fields[0].Result);
            Assert.AreEqual(string.Empty, doc.Range.Fields[1].Result);
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
            //ExFor:Range.Replace(Regex, String)
            //ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("I decided to get the curtains in gray, ideal for the grey-accented room.");

            doc.Range.Replace(new Regex("gr(a|e)y"), "lavender");

            Assert.AreEqual("I decided to get the curtains in lavender, ideal for the lavender-accented room.", doc.GetText().Trim());
            //ExEnd
        }

        //ExStart
        //ExFor:FindReplaceOptions.ReplacingCallback
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:ReplacingArgs.Replacement
        //ExFor:IReplacingCallback
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
        [Test] //ExSkip
        public void ReplaceWithCallback()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Our new location in New York City is opening tomorrow. " +
                            "Hope to see all our NYC-based customers at the opening!");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set a callback that tracks any replacements that the "Replace" method will make.
            TextFindAndReplacementLogger logger = new TextFindAndReplacementLogger();
            options.ReplacingCallback = logger;

            doc.Range.Replace(new Regex("New York City|NYC"), "Washington", options);
            
            Assert.AreEqual("Our new location in (Old value:\"New York City\") Washington is opening tomorrow. " +
                            "Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!", doc.GetText().Trim());

            Assert.AreEqual("\"New York City\" converted to \"Washington\" 20 characters into a Run node.\r\n" +
                            "\"NYC\" converted to \"Washington\" 42 characters into a Run node.", logger.GetLog().Trim());
        }

        /// <summary>
        /// Maintains a log of every text replacement done by a find-and-replace operation
        /// and notes the original matched text's value.
        /// </summary>
        private class TextFindAndReplacementLogger : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                mLog.AppendLine($"\"{args.Match.Value}\" converted to \"{args.Replacement}\" " +
                                $"{args.MatchOffset} characters into a {args.MatchNode.NodeType} node.");
                
                args.Replacement = $"(Old value:\"{args.Match.Value}\") {args.Replacement}";
                return ReplaceAction.Replace;
            }

            public string GetLog()
            {
                return mLog.ToString();
            }

            private readonly StringBuilder mLog = new StringBuilder();
        }
        //ExEnd

        //ExStart
        //ExFor:FindReplaceOptions.ApplyFont
        //ExFor:FindReplaceOptions.ReplacingCallback
        //ExFor:ReplacingArgs.GroupIndex
        //ExFor:ReplacingArgs.GroupName
        //ExFor:ReplacingArgs.Match
        //ExFor:ReplacingArgs.MatchOffset
        //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
        [Test] //ExSkip
        public void ConvertNumbersToHexadecimal()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n" +
                            "123, 456, 789 and 17379.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
            options.ApplyFont.HighlightColor = Color.LightGray;
            
            NumberHexer numberHexer = new NumberHexer();
            options.ReplacingCallback = numberHexer;

            int replacementCount = doc.Range.Replace(new Regex("[0-9]+"), "", options);

            Console.WriteLine(numberHexer.GetLog());

            Assert.AreEqual(4, replacementCount);
            Assert.AreEqual("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r" +
                            "0x7B, 0x1C8, 0x315 and 0x43E3.", doc.GetText().Trim());
            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Run, true).OfType<Run>()
                    .Count(r => r.Font.HighlightColor.ToArgb() == Color.LightGray.ToArgb()));
        }

        /// <summary>
        /// Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
        /// Maintains a log of every replacement.
        /// </summary>
        private class NumberHexer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                mCurrentReplacementNumber++;
                
                int number = Convert.ToInt32(args.Match.Value);
                
                args.Replacement = $"0x{number:X}";

                mLog.AppendLine($"Match #{mCurrentReplacementNumber}");
                mLog.AppendLine($"\tOriginal value:\t{args.Match.Value}");
                mLog.AppendLine($"\tReplacement:\t{args.Replacement}");
                mLog.AppendLine($"\tOffset in parent {args.MatchNode.NodeType} node:\t{args.MatchOffset}");

                mLog.AppendLine(string.IsNullOrEmpty(args.GroupName)
                    ? $"\tGroup index:\t{args.GroupIndex}"
                    : $"\tGroup name:\t{args.GroupName}");

                return ReplaceAction.Replace;
            }

            public string GetLog()
            {
                return mLog.ToString();
            }

            private int mCurrentReplacementNumber;
            private readonly StringBuilder mLog = new StringBuilder();
        }
        //ExEnd

        [Test]
        public void ApplyParagraphFormat()
        {
            //ExStart
            //ExFor:FindReplaceOptions.ApplyParagraphFormat
            //ExFor:Range.Replace(String, String)
            //ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
            builder.Writeln("This one will not!");
            builder.Write("This one also will.");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[0].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[1].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[2].ParagraphFormat.Alignment);

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
            // that contains a match that the find-and-replace operation finds.
            options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Right;

            // Replace every full stop that is right before a paragraph break with an exclamation point.
            int count = doc.Range.Replace(".&p", "!&p", options);

            Assert.AreEqual(2, count);
            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[0].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[1].ParagraphFormat.Alignment);
            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[2].ParagraphFormat.Alignment);
            Assert.AreEqual("Every paragraph that ends with a full stop like this one will be right aligned!\r" +
                            "This one will not!\r" +
                            "This one also will!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void DeleteSelection()
        {
            //ExStart
            //ExFor:Node.Range
            //ExFor:Range.Delete
            //ExSummary:Shows how to delete all the nodes from a range.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Add text to the first section in the document, and then add another section.
            builder.Write("Section 1. ");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Write("Section 2.");

            Assert.AreEqual("Section 1. \fSection 2.", doc.GetText().Trim());

            // Remove the first section entirely by removing all the nodes
            // within its range, including the section itself.
            doc.Sections[0].Range.Delete();

            Assert.AreEqual(1, doc.Sections.Count);
            Assert.AreEqual("Section 2.", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void RangesGetText()
        {
            //ExStart
            //ExFor:Range
            //ExFor:Range.Text
            //ExSummary:Shows how to get the text contents of all the nodes that a range covers.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            Assert.AreEqual("Hello world!", doc.Range.Text.Trim());
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        //ExStart
        //ExFor:FindReplaceOptions.UseLegacyOrder
        //ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
        public void UseLegacyOrder(bool useLegacyOrder)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three runs which we can search for using a regex pattern.
            // Place one of those runs inside a text box.
            builder.Writeln("[tag 1]");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            builder.Writeln("[tag 2]");
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("[tag 3]");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Assign a custom callback to the "ReplacingCallback" property.
            TextReplacementTracker callback = new TextReplacementTracker();
            options.ReplacingCallback = callback;

            // If we set the "UseLegacyOrder" property to "true", the
            // find-and-replace operation will go through all the runs outside of a text box
            // before going through the ones inside a text box.
            // If we set the "UseLegacyOrder" property to "false", the
            // find-and-replace operation will go over all the runs in a range in sequential order.
            options.UseLegacyOrder = useLegacyOrder;

            doc.Range.Replace(new Regex(@"\[tag \d*\]"), "", options);

            Assert.AreEqual(useLegacyOrder ?
                new List<string> { "[tag 1]", "[tag 3]", "[tag 2]" } :
                new List<string> { "[tag 1]", "[tag 2]", "[tag 3]" }, callback.Matches);
        }

        /// <summary>
        /// Records the order of all matches that occur during a find-and-replace operation.
        /// </summary>
        private class TextReplacementTracker : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Matches.Add(e.Match.Value);
                return ReplaceAction.Replace;
            }

            public List<string> Matches { get; } = new List<string>();
        }
        //ExEnd

        [TestCase(false)]
        [TestCase(true)]
        public void UseSubstitutions(bool useSubstitutions)
        {
            //ExStart
            //ExFor:FindReplaceOptions.UseSubstitutions
            //ExSummary:Shows how to replace the text with substitutions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("John sold a car to Paul.");
            builder.Writeln("Jane sold a house to Joe.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "UseSubstitutions" property to "true" to get
            // the find-and-replace operation to recognize substitution elements.
            // Set the "UseSubstitutions" property to "false" to ignore substitution elements.
            options.UseSubstitutions = useSubstitutions;

            Regex regex = new Regex(@"([A-z]+) sold a ([A-z]+) to ([A-z]+)");
            doc.Range.Replace(regex, @"$3 bought a $2 from $1", options);

            Assert.AreEqual(
                useSubstitutions
                    ? "Paul bought a car from John.\rJoe bought a house from Jane."
                    : "$3 bought a $2 from $1.\r$3 bought a $2 from $1.", doc.GetText().Trim());
            //ExEnd
        }

        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:IReplacingCallback
        //ExFor:ReplaceAction
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:ReplacingArgs.MatchNode
        //ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
        [Test] //ExSkip
        public void InsertDocumentAtReplace()
        {
            Document mainDoc = new Document(MyDir + "Document insertion destination.docx");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();
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

                // Insert a document after the paragraph containing the matched text.
                Paragraph para = (Paragraph)args.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                // Remove the paragraph with the matched text.
                para.Remove();

                return ReplaceAction.Skip;
            }
        }

        /// <summary>
        /// Inserts all the nodes of another document after a paragraph or table.
        /// </summary>
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType.Equals(NodeType.Paragraph) || insertionDestination.NodeType.Equals(NodeType.Table))
            {
                CompositeNode dstStory = insertionDestination.ParentNode;

                NodeImporter importer =
                    new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

                foreach (Section srcSection in docToInsert.Sections.OfType<Section>())
                    foreach (Node srcNode in srcSection.Body)
                    {
                        // Skip the node if it is the last empty paragraph in a section.
                        if (srcNode.NodeType.Equals(NodeType.Paragraph))
                        {
                            Paragraph para = (Paragraph)srcNode;
                            if (para.IsEndOfSection && !para.HasChildNodes)
                                continue;
                        }

                        Node newNode = importer.ImportNode(srcNode, true);

                        dstStory.InsertAfter(newNode, insertionDestination);
                        insertionDestination = newNode;
                    }
            }
            else
            {
                throw new ArgumentException("The destination node must be either a paragraph or table.");
            }
        }
        //ExEnd

        private static void TestInsertDocumentAtReplace(Document doc)
        {
            Assert.AreEqual("1) At text that can be identified by regex:\rHello World!\r" +
                            "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                            "3) At a bookmark:", doc.FirstSection.Body.GetText().Trim());
        }

        //ExStart
        //ExFor:FindReplaceOptions.Direction
        //ExFor:FindReplaceDirection
        //ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
        [TestCase(FindReplaceDirection.Backward)] //ExSkip
        [TestCase(FindReplaceDirection.Forward)] //ExSkip
        public void Direction(FindReplaceDirection findReplaceDirection)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three runs which we can search for using a regex pattern.
            // Place one of those runs inside a text box.
            builder.Writeln("Match 1.");
            builder.Writeln("Match 2.");
            builder.Writeln("Match 3.");
            builder.Writeln("Match 4.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Assign a custom callback to the "ReplacingCallback" property.
            TextReplacementRecorder callback = new TextReplacementRecorder();
            options.ReplacingCallback = callback;

            // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
            // operation to start from the end of the range, and traverse back to the beginning.
            // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
            // operation to start from the beginning of the range, and traverse to the end.
            options.Direction = findReplaceDirection;

            doc.Range.Replace(new Regex(@"Match \d*"), "Replacement", options);

            Assert.AreEqual("Replacement.\r" +
                            "Replacement.\r" +
                            "Replacement.\r" +
                            "Replacement.", doc.GetText().Trim());

            switch (findReplaceDirection)
            {
                case FindReplaceDirection.Forward:
                    Assert.AreEqual(new[] { "Match 1", "Match 2", "Match 3", "Match 4" }, callback.Matches);
                    break;
                case FindReplaceDirection.Backward:
                    Assert.AreEqual(new[] { "Match 4", "Match 3", "Match 2", "Match 1" }, callback.Matches);
                    break;
            }
        }

        /// <summary>
        /// Records all matches that occur during a find-and-replace operation in the order that they take place.
        /// </summary>
        private class TextReplacementRecorder : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Matches.Add(e.Match.Value);
                return ReplaceAction.Replace;
            }

            public List<string> Matches { get; } = new List<string>();
        }
        //ExEnd
    }
}