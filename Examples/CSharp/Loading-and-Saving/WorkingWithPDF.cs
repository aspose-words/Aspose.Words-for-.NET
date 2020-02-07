using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class WorkingWithPDF
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            LoadPDF(dataDir);
        }

        public static void LoadPDF(string dataDir)
        {
            //ExStart:LoadPDF
            Document doc = new Document(dataDir + "Document.pdf");
            
            dataDir = dataDir + "Document_out.pdf";
            doc.Save(dataDir);
            //ExEnd:LoadPDF
            Console.WriteLine("\nDocument saved.\nFile saved at " + dataDir);
        }
    }
}
