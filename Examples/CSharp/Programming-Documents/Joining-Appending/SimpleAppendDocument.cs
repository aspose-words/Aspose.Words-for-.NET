using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class SimpleAppendDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document to the destination document using no extra options.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(dataDir + "TestFile.SimpleAppendDocument Out.docx");

            Console.WriteLine("\nSimple document append.\nFile saved at " + dataDir + "TestFile.SimpleAppendDocument Out.docx");
        }
    }
}
