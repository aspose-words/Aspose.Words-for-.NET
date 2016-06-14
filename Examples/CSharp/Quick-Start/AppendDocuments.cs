
using System.IO;

using Aspose.Words;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class AppendDocuments
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "TestFile.Destination.doc";
            // Load the destination and source documents from disk.
            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document to the destination document while keeping the original formatting of the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);

            Console.WriteLine("\nDocument appended successfully.\nFile saved at " + dataDir);
        }
    }
}
