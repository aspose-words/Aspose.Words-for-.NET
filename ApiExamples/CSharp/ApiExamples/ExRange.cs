// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRange : ApiExampleBase
    {
        #region Replace 

        [Test]
        public void ReplaceSimple()
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExFor:FindReplaceOptions
            //ExFor:FindReplaceOptions.MatchCase
            //ExFor:FindReplaceOptions.FindWholeWordsOnly
            //ExSummary:Simple find and replace operation.
            // Open the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");

            // Check the document contains what we are about to test.
            Console.WriteLine(doc.FirstSection.Body.Paragraphs[0].GetText());

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = false;

            // Replace the text in the document.
            doc.Range.Replace("_CustomerName_", "James Bond", options);

            // Save the modified document.
            doc.Save(ArtifactsDir + "Range.ReplaceSimple.docx");
            //ExEnd

            Assert.AreEqual("Hello James Bond,\r\x000c", doc.GetText());
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

            doc.Save(ArtifactsDir + "ReplaceWithString.docx");
        }

        [Test]
        public void ReplaceWithRegex()
        {
            //ExStart
            //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
            Document doc = new Document(MyDir + "Document.doc");

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = false;

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);
            //ExEnd

            doc.Save(ArtifactsDir + "ReplaceWithRegex.docx");
        }

        // Note: Need more info from dev.
        [Test]
        public void ReplaceWithoutPreserveMetaCharacters()
        {
            const string text = "some text";
            const string replaceWithText = "&ldquo;";

            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write(text);

            FindReplaceOptions options = new FindReplaceOptions();
            options.PreserveMetaCharacters = false;

            doc.Range.Replace(text, replaceWithText, options);

            Assert.AreEqual("\vdquo;\f", doc.GetText());
        }

        [Test]
        public void FindAndReplaceWithPreserveMetaCharacters()
        {
            //ExStart
            //ExFor:FindReplaceOptions.PreserveMetaCharacters
            //ExSummary:Shows how to preserved meta-characters that beginning with "&".
            Document doc = new Document(MyDir + "Range.FindAndReplaceWithPreserveMetaCharacters.docx");

            FindReplaceOptions options = new FindReplaceOptions();
            options.FindWholeWordsOnly = true;
            options.PreserveMetaCharacters = true;

            doc.Range.Replace("sad", "&ldquo; some text &rdquo;", options);
            //ExEnd

            doc.Save(ArtifactsDir + "Range.FindAndReplaceWithMetacharacters.docx");
        }

        [Test]
        public void ReplaceWithInsertHtml()
        {
            //ExStart
            //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
            //ExFor:ReplacingArgs.Replacement
            //ExFor:IReplacingCallback
            //ExFor:IReplacingCallback.Replacing
            //ExFor:ReplacingArgs
            //ExFor:DocumentBuilder.InsertHtml(String)
            //ExSummary:Replaces text specified with regular expression with HTML.
            // Open the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello <CustomerName>,");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithHtmlEvaluator(options);

            doc.Range.Replace(new Regex(@" <CustomerName>,"), String.Empty, options);

            // Save the modified document.
            doc.Save(ArtifactsDir + "Range.ReplaceWithInsertHtml.doc");

            Assert.AreEqual("James Bond, Hello\r\x000c", doc.GetText()); //ExSkip
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
        //ExEnd

        [Test]
        public void ReplaceNumbersAsHex()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Write(
                "There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

            FindReplaceOptions options = new FindReplaceOptions();

            // Highlight newly inserted content.
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback = new NumberHexer();

            int count = doc.Range.Replace(new Regex("[0-9]+"), "", options);
        }

        // Customer defined callback.
        private class NumberHexer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                // Parse numbers.
                int number = Convert.ToInt32(args.Match.Value);

                // And write it as HEX.
                args.Replacement = String.Format("0x{0:X}", number);

                return ReplaceAction.Replace;
            }
        }

        #endregion

        [Test]
        public void DeleteSelection()
        {
            //ExStart
            //ExFor:Node.Range
            //ExFor:Range.Delete
            //ExSummary:Shows how to delete all characters of a range.
            // Open Word document.
            Document doc = new Document(MyDir + "Range.DeleteSection.doc");

            // The document contains two sections. Each section has a paragraph of text.
            Console.WriteLine(doc.GetText());

            // Delete the first section from the document.
            doc.Sections[0].Range.Delete();

            // Check the first section was deleted by looking at the text of the whole document again.
            Console.WriteLine(doc.GetText());
            //ExEnd

            Assert.AreEqual("Hello2\x000c", doc.GetText());
        }

        [Test]
        public void RangesGetText()
        {
            //ExStart
            //ExFor:Range
            //ExFor:Range.Text
            //ExId:RangesGetText
            //ExSummary:Shows how to get plain, unformatted text of a range.
            Document doc = new Document(MyDir + "Document.doc");
            String text = doc.Range.Text;
            //ExEnd
        }
    }
}