
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithRevisions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            AcceptRevisions(dataDir);
            GetRevisionTypes(dataDir);
        }

        private static void AcceptRevisions(string dataDir)
        {
            // ExStart:AcceptAllRevisions
            Document doc = new Document(dataDir + "Document.doc");

            // Start tracking and make some revisions.
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document.
            doc.AcceptAllRevisions();

            dataDir = dataDir + "Document.AcceptedRevisions_out.doc";
            doc.Save(dataDir);
            // ExEnd:AcceptAllRevisions
            Console.WriteLine("\nAll revisions accepted.\nFile saved at " + dataDir);
        }

        private static void GetRevisionTypes(string dataDir)
        {
            // ExStart:GetRevisionTypes
            Document doc = new Document(dataDir + "Revisions.docx");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            // ExEnd:GetRevisionTypes
        }
    }
}
