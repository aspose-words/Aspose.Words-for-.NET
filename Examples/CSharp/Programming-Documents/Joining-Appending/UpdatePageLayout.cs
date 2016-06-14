using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class UpdatePageLayout
    {
        public static void Run()
        {
            //ExStart:UpdatePageLayout
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";

            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document 
            // is appended then any changes made after will not be reflected in the rendered output.
            dstDoc.UpdatePageLayout();

            // Join the documents.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
            // If not called again the appended document will not appear in the output of the next rendering.
            dstDoc.UpdatePageLayout();

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the joined document to PDF.
            dstDoc.Save(dataDir);
            //ExEnd:UpdatePageLayout
            Console.WriteLine("\nDocument appended successfully with updated page layout after appending the document.\nFile saved at " + dataDir);
        }
    }
}
