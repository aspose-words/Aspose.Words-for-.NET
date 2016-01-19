
using System.IO;

using Aspose.Words;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Quick_Start
{
    class AppendDocuments
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            // Load the destination and source documents from disk.
            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document to the destination document while keeping the original formatting of the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(dataDir + "TestFile Out.docx");

            Console.WriteLine("\nDocument appended successfully.\nFile saved at " + dataDir + "TestFile Out.docx");
        }
    }
}
