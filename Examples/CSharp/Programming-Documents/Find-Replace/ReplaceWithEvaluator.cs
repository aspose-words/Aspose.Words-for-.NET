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
            //ExStart:ReplaceWithEvaluator
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();
            Document doc = new Document(dataDir + "Range.ReplaceWithEvaluator.doc");

            FindReplaceOptions options = new FindReplaceOptions();   
            options.ReplacingCallback = new MyReplaceEvaluator();

            doc.Range.Replace(new Regex("[s|m]ad"), "", options);

            dataDir = dataDir + "Range.ReplaceWithEvaluator_out_.doc";
            doc.Save(dataDir);
            //ExEnd:ReplaceWithEvaluator
            Console.WriteLine("\nText replaced successfully with evaluator.\nFile saved at " + dataDir);
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
                e.Replacement = e.Match.ToString() + mMatchNumber.ToString();
                mMatchNumber++;
                return ReplaceAction.Replace;
            }

            private int mMatchNumber;
        }
        //ExEnd:MyReplaceEvaluator        
    }
}
