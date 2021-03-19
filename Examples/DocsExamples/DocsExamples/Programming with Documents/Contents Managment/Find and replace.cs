using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Managment
{
    internal class FindAndReplace : DocsExamplesBase
    {
        [Test]
        public void SimpleFindReplace()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");
            Console.WriteLine("Original document text: " + doc.Range.Text);

            doc.Range.Replace("_CustomerName_", "James Bond", new FindReplaceOptions(FindReplaceDirection.Forward));

            Console.WriteLine("Document text after replace: " + doc.Range.Text);

            // Save the modified document
            doc.Save(ArtifactsDir + "FindAndReplace.SimpleFindReplace.docx");
        }

        [Test]
        public void FindAndHighlight()
        {
            //ExStart:FindAndHighlight
            Document doc = new Document(MyDir + "Find and highlight.docx");

            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new ReplaceEvaluatorFindAndHighlight(), Direction = FindReplaceDirection.Backward
            };

            Regex regex = new Regex("your document", RegexOptions.IgnoreCase);
            doc.Range.Replace(regex, "", options);

            doc.Save(ArtifactsDir + "FindAndReplace.FindAndHighlight.docx");
            //ExEnd:FindAndHighlight
        }

        //ExStart:ReplaceEvaluatorFindAndHighlight
        private class ReplaceEvaluatorFindAndHighlight : IReplacingCallback
        {
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // in this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run) currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting.
                List<Run> runs = new List<Run>();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    remainingLength > 0 &&
                    currentNode != null &&
                    currentNode.GetText().Length <= remainingLength)
                {
                    runs.Add((Run) currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    // Select the next Run node.
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while (currentNode != null && currentNode.NodeType != NodeType.Run);
                }

                // Split the last run that contains the match if there is any text left.
                if (currentNode != null && remainingLength > 0)
                {
                    SplitRun((Run) currentNode, remainingLength);
                    runs.Add((Run) currentNode);
                }

                // Now highlight all runs in the sequence.
                foreach (Run run in runs)
                    run.Font.HighlightColor = Color.Yellow;

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }
        }
        //ExEnd:ReplaceEvaluatorFindAndHighlight

        //ExStart:SplitRun
        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run) run.Clone(true);
            afterRun.Text = run.Text.Substring(position);

            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            
            return afterRun;
        }
        //ExEnd:SplitRun

        [Test]
        public void MetaCharactersInSearchPattern()
        {
            /* meta-characters
            &p - paragraph break
            &b - section break
            &m - page break
            &l - manual line break
            */

            //ExStart:MetaCharactersInSearchPattern
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("This is Line 1");
            builder.Writeln("This is Line 2");

            doc.Range.Replace("This is Line 1&pThis is Line 2", "This is replaced line");

            builder.MoveToDocumentEnd();
            builder.Write("This is Line 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is Line 2");

            doc.Range.Replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.");

            doc.Save(ArtifactsDir + "FindAndReplace.MetaCharactersInSearchPattern.docx");
            //ExEnd:MetaCharactersInSearchPattern
        }

        [Test]
        public void ReplaceTextContainingMetaCharacters()
        {
            //ExStart:ReplaceTextContainingMetaCharacters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("First section");
            builder.Writeln("  1st paragraph");
            builder.Writeln("  2nd paragraph");
            builder.Writeln("{insert-section}");
            builder.Writeln("Second section");
            builder.Writeln("  1st paragraph");

            FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
            findReplaceOptions.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Double each paragraph break after word "section", add kind of underline and make it centered.
            int count = doc.Range.Replace("section&p", "section&p----------------------&p", findReplaceOptions);

            // Insert section break instead of custom text tag.
            count = doc.Range.Replace("{insert-section}", "&b", findReplaceOptions);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceTextContainingMetaCharacters.docx");
            //ExEnd:ReplaceTextContainingMetaCharacters
        }

        [Test]
        public void IgnoreTextInsideFields()
        {
            //ExStart:IgnoreTextInsideFields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert field with text inside.
            builder.InsertField("INCLUDETEXT", "Text in field");
            
            FindReplaceOptions options = new FindReplaceOptions { IgnoreFields = true };
            
            Regex regex = new Regex("e");
            doc.Range.Replace(regex, "*", options);
            
            Console.WriteLine(doc.GetText());

            options.IgnoreFields = false;
            doc.Range.Replace(regex, "*", options);
            
            Console.WriteLine(doc.GetText());
            //ExEnd:IgnoreTextInsideFields
        }

        [Test]
        public void IgnoreTextInsideDeleteRevisions()
        {
            //ExStart:IgnoreTextInsideDeleteRevisions
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert non-revised text.
            builder.Writeln("Deleted");
            builder.Write("Text");

            // Remove first paragraph with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            doc.FirstSection.Body.FirstParagraph.Remove();
            doc.StopTrackRevisions();

            FindReplaceOptions options = new FindReplaceOptions { IgnoreDeleted = true };

            Regex regex = new Regex("e");
            doc.Range.Replace(regex, "*", options);

            Console.WriteLine(doc.GetText());

            options.IgnoreDeleted = false;
            doc.Range.Replace(regex, "*", options);

            Console.WriteLine(doc.GetText());
            //ExEnd:IgnoreTextInsideDeleteRevisions
        }

        [Test]
        public void IgnoreTextInsideInsertRevisions()
        {
            //ExStart:IgnoreTextInsideInsertRevisions
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with tracking revisions.
            doc.StartTrackRevisions("author", DateTime.Now);
            builder.Writeln("Inserted");
            doc.StopTrackRevisions();

            // Insert non-revised text.
            builder.Write("Text");

            FindReplaceOptions options = new FindReplaceOptions { IgnoreInserted = true };

            Regex regex = new Regex("e");
            doc.Range.Replace(regex, "*", options);
            
            Console.WriteLine(doc.GetText());

            options.IgnoreInserted = false;
            doc.Range.Replace(regex, "*", options);
            
            Console.WriteLine(doc.GetText());
            //ExEnd:IgnoreTextInsideInsertRevisions
        }

        [Test]
        public void ReplaceHtmlTextWithMetaCharacters()
        {
            //ExStart:ReplaceHtmlTextWithMetaCharacters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("{PLACEHOLDER}");

            FindReplaceOptions findReplaceOptions = new FindReplaceOptions { ReplacingCallback = new FindAndInsertHtml() };

            doc.Range.Replace("{PLACEHOLDER}", "<p>&ldquo;Some Text&rdquo;</p>", findReplaceOptions);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceHtmlTextWithMetaCharacters.docx");
            //ExEnd:ReplaceHtmlTextWithMetaCharacters
        }

        //ExStart:ReplaceHtmlFindAndInsertHtml
        public sealed class FindAndInsertHtml : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Node currentNode = e.MatchNode;

                DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);
                builder.MoveTo(currentNode);
                builder.InsertHtml(e.Replacement);

                currentNode.Remove();

                return ReplaceAction.Skip;
            }
        }
        //ExEnd:ReplaceHtmlFindAndInsertHtml

        [Test]
        public void ReplaceTextInFooter()
        {
            //ExStart:ReplaceTextInFooter
            Document doc = new Document(MyDir + "Footer.docx");

            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];

            FindReplaceOptions options = new FindReplaceOptions { MatchCase = false, FindWholeWordsOnly = false };

            footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceTextInFooter.docx");
            //ExEnd:ReplaceTextInFooter
        }

        [Test]
        //ExStart:ShowChangesForHeaderAndFooterOrders
        public void ShowChangesForHeaderAndFooterOrders()
        {
            ReplaceLog logger = new ReplaceLog();
            
            Document doc = new Document(MyDir + "Footer.docx");
            Section firstPageSection = doc.FirstSection;
            
            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = logger };

            doc.Range.Replace(new Regex("(header|footer)"), "", options);
            
            doc.Save(ArtifactsDir + "FindAndReplace.ShowChangesForHeaderAndFooterOrders.docx");

            logger.ClearText();

            firstPageSection.PageSetup.DifferentFirstPageHeaderFooter = false;

            doc.Range.Replace(new Regex("(header|footer)"), "", options);
        }

        private class ReplaceLog : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                mTextBuilder.AppendLine(args.MatchNode.GetText());
                return ReplaceAction.Skip;
            }

            internal void ClearText()
            {
                mTextBuilder.Clear();
            }

            private readonly StringBuilder mTextBuilder = new StringBuilder();
        }
        //ExEnd:ShowChangesForHeaderAndFooterOrders

        [Test]
        public void ReplaceTextWithField()
        {
            Document doc = new Document(MyDir + "Replace text with fields.docx");

            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new ReplaceTextWithFieldHandler(FieldType.FieldMergeField)
            };

            doc.Range.Replace(new Regex(@"PlaceHolder(\d+)"), "", options);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceTextWithField.docx");
        }


        public class ReplaceTextWithFieldHandler : IReplacingCallback
        {
            public ReplaceTextWithFieldHandler(FieldType type)
            {
                mFieldType = type;
            }

            public ReplaceAction Replacing(ReplacingArgs args)
            {
                List<Run> runs = FindAndSplitMatchRuns(args);

                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo(runs[runs.Count - 1]);

                // Calculate the field's name from the FieldType enumeration by removing
                // the first instance of "Field" from the text. This works for almost all of the field types.
                string fieldName = mFieldType.ToString().ToUpper().Substring(5);

                // Insert the field into the document using the specified field type and the matched text as the field name.
                // If the fields you are inserting do not require this extra parameter, it can be removed from the string below.
                builder.InsertField($"{fieldName} {args.Match.Groups[0]}");

                foreach (Run run in runs)
                    run.Remove();

                return ReplaceAction.Skip;
            }

            /// <summary>
            /// Finds and splits the match runs and returns them in an List.
            /// </summary>
            public List<Run> FindAndSplitMatchRuns(ReplacingArgs args)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = args.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run.
                if (args.MatchOffset > 0)
                    currentNode = SplitRun((Run) currentNode, args.MatchOffset);

                // This array is used to store all nodes of the match for further removing.
                List<Run> runs = new List<Run>();

                // Find all runs that contain parts of the match string.
                int remainingLength = args.Match.Value.Length;
                while (
                    remainingLength > 0 &&
                    currentNode != null &&
                    currentNode.GetText().Length <= remainingLength)
                {
                    runs.Add((Run) currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    do
                    {
                        currentNode = currentNode.NextSibling;
                    } while (currentNode != null && currentNode.NodeType != NodeType.Run);
                }

                // Split the last run that contains the match if there is any text left.
                if (currentNode != null && remainingLength > 0)
                {
                    SplitRun((Run) currentNode, remainingLength);
                    runs.Add((Run) currentNode);
                }

                return runs;
            }

            /// <summary>
            /// Splits text of the specified run into two runs.
            /// Inserts the new run just after the specified run.
            /// </summary>
            private Run SplitRun(Run run, int position)
            {
                Run afterRun = (Run) run.Clone(true);
                
                afterRun.Text = run.Text.Substring(position);
                run.Text = run.Text.Substring(0, position);
                
                run.ParentNode.InsertAfter(afterRun, run);
                
                return afterRun;
            }

            private readonly FieldType mFieldType;
        }

        [Test]
        public void ReplaceWithEvaluator()
        {
            //ExStart:ReplaceWithEvaluator
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = new MyReplaceEvaluator() };

            doc.Range.Replace(new Regex("[s|m]ad"), "", options);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceWithEvaluator.docx");
            //ExEnd:ReplaceWithEvaluator
        }

        //ExStart:MyReplaceEvaluator
        private class MyReplaceEvaluator : IReplacingCallback
        {
            /// <summary>
            /// This is called during a replace operation each time a match is found.
            /// This method appends a number to the match string and returns it as a replacement string.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                e.Replacement = e.Match + mMatchNumber.ToString();
                mMatchNumber++;
                
                return ReplaceAction.Replace;
            }

            private int mMatchNumber;
        }
        //ExEnd:MyReplaceEvaluator

        [Test]
        //ExStart:ReplaceWithHtml
        public void ReplaceWithHtml()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello <CustomerName>,");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithHtmlEvaluator(options);

            doc.Range.Replace(new Regex(@" <CustomerName>,"), string.Empty, options);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceWithHtml.docx");
        }

        private class ReplaceWithHtmlEvaluator : IReplacingCallback
        {
            internal ReplaceWithHtmlEvaluator(FindReplaceOptions options)
            {
                mOptions = options;
            }

            /// <summary>
            /// NOTE: This is a simplistic method that will only work well when the match
            /// starts at the beginning of a run.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // Replace '<CustomerName>' text with a red bold name.
                builder.InsertHtml("<b><font color='red'>James Bond, </font></b>");
                args.Replacement = "";

                return ReplaceAction.Replace;
            }

            private readonly FindReplaceOptions mOptions;
        }
        //ExEnd:ReplaceWithHtml

        [Test]
        public void ReplaceWithRegex()
        {
            //ExStart:ReplaceWithRegex
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            FindReplaceOptions options = new FindReplaceOptions();

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceWithRegex.docx");
            //ExEnd:ReplaceWithRegex
        }
        
        [Test]
        public void RecognizeAndSubstitutionsWithinReplacementPatterns()
        {
            //ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Jason give money to Paul.");

            Regex regex = new Regex(@"([A-z]+) give money to ([A-z]+)");

            FindReplaceOptions options = new FindReplaceOptions { UseSubstitutions = true };

            doc.Range.Replace(regex, @"$2 take money from $1", options);
            //ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
        }

        [Test]
        public void ReplaceWithString()
        {
            //ExStart:ReplaceWithString
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("sad mad bad");

            doc.Range.Replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceWithString.docx");
            //ExEnd:ReplaceWithString
        }

        [Test]
        //ExStart:UsingLegacyOrder
        public void UsingLegacyOrder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("[tag 1]");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            builder.Writeln("[tag 3]");

            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("[tag 2]");

            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new ReplacingCallback(), UseLegacyOrder = true
            };

            doc.Range.Replace(new Regex(@"\[(.*?)\]"), "", options);

            doc.Save(ArtifactsDir + "FindAndReplace.UsingLegacyOrder.docx");
        }

        private class ReplacingCallback : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Console.Write(e.Match.Value);
                return ReplaceAction.Replace;
            }
        }
        //ExEnd:UsingLegacyOrder

        [Test]
        public void ReplaceTextInTable()
        {
            //ExStart:ReplaceText
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            table.Range.Replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.Forward));
            table.LastRow.LastCell.Range.Replace("50", "20", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "FindAndReplace.ReplaceTextInTable.docx");
            //ExEnd:ReplaceText
        }
    }
}
