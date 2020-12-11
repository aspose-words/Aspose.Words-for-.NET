using System;
using System.Text.RegularExpressions;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Replacing;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithEvaluator
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            ReplaceUsingPattern(dataDir);
        }

        public static void ReplaceUsingPattern(string dataDir)
        {
            // ExStart:ReplaceWithEvaluator
            // The path to the documents directory.
            
            Document doc = new Document(dataDir + "Range.ReplaceWithEvaluator.doc");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new MyReplaceEvaluator();

            doc.Range.Replace(new Regex("[s|m]ad"), "", options);

            dataDir = dataDir + "Range.ReplaceWithEvaluator_out.doc";
            doc.Save(dataDir);
            // ExEnd:ReplaceWithEvaluator
            Console.WriteLine("\nText replaced successfully with evaluator.\nFile saved at " + dataDir);
        }

        // ExStart:MyReplaceEvaluator
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
        // ExEnd:MyReplaceEvaluator
    }
}
