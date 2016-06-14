using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Ranges
{
    class RangesDeleteText
    {
        public static void Run()
        {
            //ExStart:RangesDeleteText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithRanges();

            Document doc = new Document(dataDir + "Document.doc");
            doc.Sections[0].Range.Delete();
            //ExEnd:RangesDeleteText
            Console.WriteLine("\nAll characters of a range deleted successfully.");
        }
    }
}
