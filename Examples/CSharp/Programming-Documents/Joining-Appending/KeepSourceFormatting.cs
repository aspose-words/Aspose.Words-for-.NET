using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class KeepSourceFormatting
    {
        public static void Run()
        {
            //ExStart:KeepSourceFormatting
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";
            // Load the documents to join.
            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Keep the formatting from the source document when appending it to the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the joined document to disk.
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:KeepSourceFormatting
            Console.WriteLine("\nDocument appended successfully with keep source formatting option.\nFile saved at " + dataDir);
        }
    }
}
