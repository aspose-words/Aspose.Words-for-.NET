using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Ranges
{
    class RangesGetText
    {
        public static void Run()
        {
            //ExStart:RangesGetText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithRanges();

            Document doc = new Document(dataDir + "Document.doc");
            string text = doc.Range.Text; 
            //ExEnd:RangesGetText
            Console.WriteLine("\nDocument have following text range " + text);
        }
    }
}
