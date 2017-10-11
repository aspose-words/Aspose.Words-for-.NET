using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveOptionsHtmlFixed
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            UseFontFromTargetMachine(dataDir);

        }

        static void UseFontFromTargetMachine(string dataDir)
        {
            // ExStart:UseFontFromTargetMachine
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (doc).doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
            options.UseTargetMachineFonts = true;

            dataDir = dataDir + "Test File_out.html";

            // Save the document to disk.
            doc.Save(dataDir, options);
            // ExEnd:UseFontFromTargetMachine
            Console.WriteLine("\nTable cloned successfully.\nFile saved at " + dataDir);
        }
    }
}
