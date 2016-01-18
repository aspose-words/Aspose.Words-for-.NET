using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class UseDestinationStyles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            // Load the documents to join.
            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document using the styles of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            // Save the joined document to disk.
            dstDoc.Save(dataDir + "TestFile.UseDestinationStyles Out.doc");

            Console.WriteLine("\nDocument appended successfully with use destination styles option.\nFile saved at " + dataDir + "TestFile.UseDestinationStyles Out.doc");
        }
    }
}
