using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_Replace
{
    class ReplaceWithHTML
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            ReplaceWithHtml(dataDir);
        }

        // ExStart:ReplaceWithHtml
        public static void ReplaceWithHtml(string dataDir)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello <CustomerName>,");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplaceWithHtmlEvaluator(options);

            doc.Range.Replace(new Regex(@" <CustomerName>,"), String.Empty, options);

            // Save the modified document. 
            doc.Save(dataDir + "Range.ReplaceWithInsertHtml.doc");
        }

        private class ReplaceWithHtmlEvaluator : IReplacingCallback
        {
            internal ReplaceWithHtmlEvaluator(FindReplaceOptions options)
            {
                mOptions = options;
            }

            //This simplistic method will only work well when the match starts at the beginning of a run. 
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // Replace '<CustomerName>' text with a red bold name.
                builder.InsertHtml("<b><font color='red'>James Bond, </font></b>"); args.Replacement = "";
                return ReplaceAction.Replace;
            }

            private readonly FindReplaceOptions mOptions;
        }
        // ExEnd:ReplaceWithHtml

        // ExStart:NumberHighlightCallback
        // Replace and Highlight Numbers.
        internal class NumberHighlightCallback : IReplacingCallback
        {
            public NumberHighlightCallback(FindReplaceOptions opt)
            {
                mOpt = opt;
            }

            public ReplaceAction Replacing(ReplacingArgs args)
            {
                // Let replacement to be the same text.
                args.Replacement = args.Match.Value;

                int val = int.Parse(args.Match.Value);

                // Apply either red or green color depending on the number value sign.
                mOpt.ApplyFont.Color = (val > 0)
                    ? Color.Green
                    : Color.Red;

                return ReplaceAction.Replace;
            }

            private readonly FindReplaceOptions mOpt;
        }
        // ExEnd:NumberHighlightCallback

        // ExStart:LineCounter
        public void LineCounter()
        {
            // Create a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add lines of text.
            builder.Writeln("This is first line");
            builder.Writeln("Second line");
            builder.Writeln("And last line");

            // Prepend each line with line number.
            FindReplaceOptions opt = new FindReplaceOptions() { ReplacingCallback = new LineCounterCallback() };
            doc.Range.Replace(new Regex("[^&p]*&p"), "", opt);

            doc.Save(@"X:\TestLineCounter.docx");
        }

        internal class LineCounterCallback : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                Debug.WriteLine(args.Match.Value);

                args.Replacement = string.Format("{0} {1}", mCounter++, args.Match.Value);
                return ReplaceAction.Replace;
            }

            private int mCounter = 1;
        }
        // ExEnd:LineCounter
    }
}
