using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class UnlinkHeadersFooters
    {
        public static void Run()
        {
            //ExStart:UnlinkHeadersFooters
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";

            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
            // from the destination document.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:UnlinkHeadersFooters
            Console.WriteLine("\nDocument appended successfully with unlinked header footers.\nFile saved at " + dataDir);
        }
    }
}
