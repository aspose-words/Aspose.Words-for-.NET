
using System.IO;

using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveToDisk
    {
        public static void Run()
        {
            //ExStart:LoadAndSave
            //ExStart:OpenDocument
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "Document.doc";
            // Load the document from the absolute path on disk.
            Document doc = new Document(dataDir + fileName);
            //ExEnd:OpenDocument
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document as DOCX document.");
            doc.Save(dataDir);
            //ExEnd:LoadAndSave
            Console.WriteLine("\nExisting document loaded and saved successfully.\nFile saved at " + dataDir);
        }
    }
}
