using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class DifferentPageSetup
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.SourcePageSetup.doc");

            // Set the source document to continue straight after the end of the destination document.
            // If some page setup settings are different then this may not work and the source document will appear 
            // on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // To ensure this does not happen when the source document has different page setup settings make sure the
            // settings are identical between the last section of the destination document.
            // If there are further continuous sections that follow on in the source document then this will need to be 
            // repeated for those sections as well.
            srcDoc.FirstSection.PageSetup.PageWidth = dstDoc.LastSection.PageSetup.PageWidth;
            srcDoc.FirstSection.PageSetup.PageHeight = dstDoc.LastSection.PageSetup.PageHeight;
            srcDoc.FirstSection.PageSetup.Orientation = dstDoc.LastSection.PageSetup.Orientation;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.DifferentPageSetup Out.doc");

            Console.WriteLine("\nDocument appended successfully with different page setup.\nFile saved at " + dataDir + "TestFile.DifferentPageSetup Out.doc");
        }
    }
}
