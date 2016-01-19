using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class JoinContinuous
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Make the document appear straight after the destination documents content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Append the source document using the original styles found in the source document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.JoinContinuous Out.doc");

            Console.WriteLine("\nDocument appended successfully with continous section start.\nFile saved at " + dataDir + "TestFile.JoinContinuous Out.doc");
        }
    }
}
