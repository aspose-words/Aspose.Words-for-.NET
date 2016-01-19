using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class KeepSourceFormatting
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            // Load the documents to join.
            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Keep the formatting from the source document when appending it to the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the joined document to disk.
            dstDoc.Save(dataDir + "TestFile.KeepSourceFormatting Out.docx");

            Console.WriteLine("\nDocument appended successfully with keep source formatting option.\nFile saved at " + dataDir + "TestFile.KeepSourceFormatting Out.docx");
        }
    }
}
