using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using System;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithRegex
    {
        public static void Run()
        {
            //ExStart:ReplaceWithRegex
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            Document doc = new Document(dataDir + "Document.doc");

            FindReplaceOptions options = new FindReplaceOptions();

            doc.Range.Replace(new Regex("[s|m]ad"), "bad", options);

            dataDir = dataDir + "ReplaceWithRegex_out_.doc";
            doc.Save(dataDir);
            //ExEnd:ReplaceWithRegex
            Console.WriteLine("\nText replaced with regex successfully.\nFile saved at " + dataDir);
        }
    }    
}
