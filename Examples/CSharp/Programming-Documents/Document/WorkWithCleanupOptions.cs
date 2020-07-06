using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkWithCleanupOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            CleanupUnusedStylesandLists(dataDir);
            CleanupDuplicateStyle(dataDir);
        }

        private static void CleanupUnusedStylesandLists(string dataDir)
        {
            // ExStart:CleanupUnusedStylesandLists
            Document doc = new Document(dataDir + "Document.doc");

            CleanupOptions cleanupoptions = new CleanupOptions();
            cleanupoptions.UnusedLists = false;
            cleanupoptions.UnusedStyles = true;

            // Cleans unused styles and lists from the document depending on given CleanupOptions. 
            doc.Cleanup(cleanupoptions);

            dataDir = dataDir + "Document.CleanupUnusedStylesandLists_out.docx";
            doc.Save(dataDir);
            // ExEnd:CleanupUnusedStylesandLists
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }

        private static void CleanupDuplicateStyle(string dataDir)
        {
            // ExStart:CleanupDuplicateStyle
            Document doc = new Document(dataDir + "Document.doc");

            CleanupOptions options = new CleanupOptions();
            options.DuplicateStyle = true;

            // Cleans duplicate styles from the document. 
            doc.Cleanup(options);

            doc.Save(dataDir + "Document.CleanupDuplicateStyle_out.docx");
            // ExEnd:CleanupDuplicateStyle
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }
    }
}
