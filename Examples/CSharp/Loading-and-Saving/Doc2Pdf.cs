
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class Doc2Pdf
    {
        public static void Run()
        {
            //ExStart:Doc2Pdf
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Template.doc");

            dataDir = dataDir + "Template_out_.pdf";

            // Save the document in PDF format.
            doc.Save(dataDir);
            //ExEnd:Doc2Pdf
            Console.WriteLine("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
        }
    }
}
