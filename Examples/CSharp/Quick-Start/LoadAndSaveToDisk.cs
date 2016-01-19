
using System.IO;

using Aspose.Words;
using System;

namespace CSharp.Quick_Start
{
    class LoadAndSaveToDisk
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "Document.doc";

            // Load the document from the absolute path on disk.
            Document doc = new Document(dataDir + fileName);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document as DOCX document.");
            doc.Save(dataDir);

            Console.WriteLine("\nExisting document loaded and saved successfully.\nFile saved at " + dataDir);
        }
    }
}
