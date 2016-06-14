using System;
using System.IO;

using Aspose.Words;
using System.Collections;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class PrependDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            string fileName = "TestFile.Destination.doc";

            Document dstDoc = new Document(dataDir + fileName);
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document to the destination document. This causes the result to have line spacing problems.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Instead prepend the content of the destination document to the start of the source document.
            // This results in the same joined document but with no line spacing issues.
            DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document
            dstDoc.Save(dataDir);

            Console.WriteLine("\nDocument prepended successfully.\nFile saved at " + dataDir);
        }

        public static void DoPrepend(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document. 
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            ArrayList sections = new ArrayList(srcDoc.Sections.ToArray());

            // Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
            sections.Reverse();

            foreach (Section srcSection in sections)
            {
                // Import the nodes from the source document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be prepended to the destination document.
                // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
                // to the original method.
                dstDoc.PrependChild(dstSection);
            }
        }
    }
}
