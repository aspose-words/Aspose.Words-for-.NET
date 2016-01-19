using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class ListKeepSourceFormatting
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.DestinationList.doc");
            Document srcDoc = new Document(dataDir + "TestFile.SourceList.doc");

            // Append the content of the document so it flows continuously.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.ListKeepSourceFormatting Out.doc");

            Console.WriteLine("\nDocument appended successfully with lists retaining source formatting.\nFile saved at " + dataDir + "TestFile.ListKeepSourceFormatting Out.doc");
        }
    }
}
