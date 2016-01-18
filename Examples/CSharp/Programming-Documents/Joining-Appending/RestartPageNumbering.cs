using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class RestartPageNumbering
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Set the appended document to appear on the next page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Restart the page numbering for the document to be appended.
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.RestartPageNumbering Out.doc");

            Console.WriteLine("\nDocument appended successfully with restart page numbering.\nFile saved at " + dataDir + "TestFile.RestartPageNumbering Out.doc");
        }
    }
}
