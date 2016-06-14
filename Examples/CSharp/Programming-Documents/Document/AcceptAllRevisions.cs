
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class AcceptAllRevisions
    {
        public static void Run()
        {
            //ExStart:AcceptAllRevisions
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document(dataDir + "Document.doc");

            // Start tracking and make some revisions.
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document.
            doc.AcceptAllRevisions();

            dataDir = dataDir + "Document.AcceptedRevisions_out_.doc";
            doc.Save(dataDir);
            //ExEnd:AcceptAllRevisions
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }        
    }
}
