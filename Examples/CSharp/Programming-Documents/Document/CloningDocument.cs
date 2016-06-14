
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CloningDocument
    {
        public static void Run()
        {
            //ExStart:CloningDocument
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // Load the document from disk.
            Document doc = new Document(dataDir + "TestFile.doc");

            Document clone = doc.Clone();

            dataDir = dataDir + "TestFile_clone_out_.doc";

            // Save the document to disk.
            clone.Save(dataDir);
            //ExEnd:CloningDocument
            Console.WriteLine("\nDocument cloned successfully.\nFile saved at " + dataDir);
        }
    }
}
