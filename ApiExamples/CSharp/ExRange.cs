// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
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
        [Test]
        public void DeleteSelection()
        {
            //ExStart
            //ExFor:Node.Range
            //ExFor:Range.Delete
            //ExSummary:Shows how to delete a section from a Word document.
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
        public void ReplaceSimple()
        {
            //ExStart
            //ExFor:Range.Replace(String,String,Boolean,Boolean)
            //ExSummary:Simple find and replace operation.
            // Open the document.
            Document doc = new Document(MyDir + "Range.ReplaceSimple.doc");

            // Check the document contains what we are about to test.
            Console.WriteLine(doc.FirstSection.Body.Paragraphs[0].GetText());

            // Replace the text in the document.
            doc.Range.Replace("_CustomerName_", "James Bond", false, false);

            // Save the modified document.
            doc.Save(MyDir + @"\Artifacts\Range.ReplaceSimple.doc");
            //ExEnd

            Assert.AreEqual("Hello James Bond,\r\x000c", doc.GetText());
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ReplaceWithInsertHtmlCaller()
        {
            this.ReplaceWithInsertHtml();
        }

        //ExStart
        //ExFor:Range.Replace(Regex,IReplacingCallback,Boolean)
        //ExFor:ReplacingArgs.Replacement
        //ExFor:IReplacingCallback
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:DocumentBuilder.InsertHtml(string)
        //ExSummary:Replaces text specified with regular expression with HTML.
        public void ReplaceWithInsertHtml()
        {
            // Open the document.
            Document doc = new Document(MyDir + "Range.ReplaceWithInsertHtml.doc");

            doc.Range.Replace(new Regex(@"<CustomerName>"), new ReplaceWithHtmlEvaluator(), false);

            // Save the modified document.
            doc.Save(MyDir + @"\Artifacts\Range.ReplaceWithInsertHtml.doc");

            Assert.AreEqual("Hello James Bond,\r\x000c", doc.GetText());  //ExSkip
        }

        private class ReplaceWithHtmlEvaluator : IReplacingCallback
        {
            /// <summary>
            /// NOTE: This is a simplistic method that will only work well when the match
            /// starts at the beginning of a run.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                DocumentBuilder builder = new DocumentBuilder((Document)e.MatchNode.Document);
                builder.MoveTo(e.MatchNode);
                // Replace '<CustomerName>' text with a red bold name.
                builder.InsertHtml("<b><font color='red'>James Bond</font></b>");

                e.Replacement = "";
                return ReplaceAction.Replace;
            }
        }
        //ExEnd

        [Test]
        public void RangesGetText()
        {
            //ExStart
            //ExFor:Range
            //ExFor:Range.Text
            //ExId:RangesGetText
            //ExSummary:Shows how to get plain, unformatted text of a range.
            Document doc = new Document(MyDir + "Document.doc");
            string text = doc.Range.Text;
            //ExEnd
        }

        [Test]
        public void ReplaceWithString()
        {
            //ExStart
            //ExFor:Range
            //ExId:RangesReplaceString
            //ExSummary:Shows how to replace all occurrences of word "sad" to "bad".
            Document doc = new Document(MyDir + "Document.doc");
            doc.Range.Replace("sad", "bad", false, true);
            //ExEnd
            doc.Save(MyDir + @"\Artifacts\ReplaceWithString.doc");
        }

        [Test]
        public void ReplaceWithRegex()
        {
            //ExStart
            //ExFor:Range.Replace(Regex, String)
            //ExId:RangesReplaceRegex
            //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
            Document doc = new Document(MyDir + "Document.doc");
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");
            //ExEnd
            doc.Save(MyDir + @"\Artifacts\ReplaceWithRegex.doc");
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ReplaceWithEvaluatorCaller()
        {
            this.ReplaceWithEvaluator();
        }

        //ExStart
        //ExFor:Range
        //ExFor:ReplacingArgs.Match
        //ExId:RangesReplaceWithReplaceEvaluator
        //ExSummary:Shows how to replace with a custom evaluator.
        public void ReplaceWithEvaluator()
        {
            Document doc = new Document(MyDir + "Range.ReplaceWithEvaluator.doc");
            doc.Range.Replace(new Regex("[s|m]ad"), new MyReplaceEvaluator(), true);
            doc.Save(MyDir + @"\Artifacts\Range.ReplaceWithEvaluator.doc");
        }

        private class MyReplaceEvaluator : IReplacingCallback
        {
            /// <summary>
            /// This is called during a replace operation each time a match is found.
            /// This method appends a number to the match string and returns it as a replacement string.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                e.Replacement = e.Match.ToString() + this.mMatchNumber.ToString();
                this.mMatchNumber++;
                return ReplaceAction.Replace;
            }

            private int mMatchNumber;
        }
        //ExEnd

        [Test]
        public void RangesDeleteText()
        {
            //ExStart
            //ExId:RangesDeleteText
            //ExSummary:Shows how to delete all characters of a range.
            Document doc = new Document(MyDir + "Document.doc");
            doc.Sections[0].Range.Delete();
            //ExEnd
        }

        /// <summary>
        /// RK This works, but the logic is so complicated that I don't want to show it to users.
        /// </summary>
        [Test]
        public void ChangeTextToHyperlinks()
        {
            Document doc = new Document(MyDir + @"Range.ChangeTextToHyperlinks.doc");

            // Create regular expression for URL search
            Regex regexUrl = new Regex(@"(?<Protocol>\w+):\/\/(?<Domain>[\w.]+\/?)\S*(?x)");

            // Run replacement, using regular expression and evaluator.
            doc.Range.Replace(regexUrl, new ChangeTextToHyperlinksEvaluator(doc), false);

            // Save updated document.
            doc.Save(MyDir + @"\Artifacts\Range.ChangeTextToHyperlinks.docx");
        }

        private class ChangeTextToHyperlinksEvaluator : IReplacingCallback
        {
            internal ChangeTextToHyperlinksEvaluator(Document doc)
            {
                this.mBuilder = new DocumentBuilder(doc);
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is the run node that contains the found text. Note that the run might contain other 
                // text apart from the URL. All the complexity below is just to handle that. I don't think there
                // is a simpler way at the moment.
                Run run = (Run)e.MatchNode;

                Paragraph para = run.ParentParagraph;

                string url = e.Match.Value;

                // We are using \xbf (inverted question mark) symbol for temporary purposes.
                // Any symbol will do that is non-special and is guaranteed not to be presented in the document.
                // The purpose is to split the matched run into two and insert a hyperlink field between them.
                para.Range.Replace(url, "\xbf", true, true);

                Run subRun = (Run)run.Clone(false);
                int pos = run.Text.IndexOf("\xbf");
                subRun.Text = subRun.Text.Substring(0, pos);
                run.Text = run.Text.Substring(pos + 1, run.Text.Length - pos - 1);

                para.ChildNodes.Insert(para.ChildNodes.IndexOf(run), subRun);

                this.mBuilder.MoveTo(run);

                // Specify font formatting for the hyperlink.
                this.mBuilder.Font.Color = Color.Blue;
                this.mBuilder.Font.Underline = Underline.Single;

                // Insert the hyperlink.
                this.mBuilder.InsertHyperlink(url, url, false);

                // Clear hyperlink formatting.
                this.mBuilder.Font.ClearFormatting();

                // Let's remove run if it is empty.
                if (run.Text.Equals(""))
                    run.Remove();

                // No replace action is necessary - we have already done what we intended to do.
                return ReplaceAction.Skip;
            }

            private readonly DocumentBuilder mBuilder;
        }
    }
}
