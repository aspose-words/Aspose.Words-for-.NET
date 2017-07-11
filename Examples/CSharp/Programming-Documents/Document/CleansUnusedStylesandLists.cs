using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CleansUnusedStylesandLists
    {
        public static void Run()
        {
            // ExStart:CleansUnusedStylesandLists
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document(dataDir + "Document.doc");
             
            CleanupOptions cleanupoptions = new CleanupOptions();
            cleanupoptions.UnusedLists = false;
            cleanupoptions.UnusedStyles = true;

            // Cleans unused styles and lists from the document depending on given CleanupOptions. 
            doc.Cleanup(cleanupoptions);

            dataDir = dataDir + "Document.Cleanup_out.docx";
            doc.Save(dataDir);
            // ExEnd:CleansUnusedStylesandLists
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }
    }
}
