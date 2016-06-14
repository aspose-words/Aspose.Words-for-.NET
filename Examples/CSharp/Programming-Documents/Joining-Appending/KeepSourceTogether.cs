using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class KeepSourceTogether
    {
        public static void Run()
        {
            //ExStart:KeepSourceTogether
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.DestinationList.doc";

            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Set the source document to appear straight after the destination document's content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Iterate through all sections in the source document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:KeepSourceTogether
            Console.WriteLine("\nDocument appended successfully while keeping the content from splitting across two pages.\nFile saved at " + dataDir);
        }
    }
}
