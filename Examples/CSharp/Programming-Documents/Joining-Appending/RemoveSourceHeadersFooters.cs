using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class RemoveSourceHeadersFooters
    {
        public static void Run()
        {
            //ExStart:RemoveSourceHeadersFooters
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";
            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Remove the headers and footers from each of the sections in the source document.
            foreach (Section section in srcDoc.Sections)
            {
                section.ClearHeadersFooters();
            }

            // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
            // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
            // document. This should set to false to avoid this behavior.
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:RemoveSourceHeadersFooters
            Console.WriteLine("\nDocument appended successfully with source header footers removed.\nFile saved at " + dataDir);
        }
    }
}
