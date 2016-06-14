using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using System.Text;
namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveAsTxt
    {
        public static void Run()
        {
            //ExStart:SaveAsTxt
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            Document doc = new Document(dataDir + "Document.doc");
            dataDir = dataDir + "Document.ConvertToTxt_out_.txt";
            doc.Save(dataDir);
            //ExEnd:SaveAsTxt
            Console.WriteLine("\nDocument saved as TXT.\nFile saved at " + dataDir);
        }
    }
}
