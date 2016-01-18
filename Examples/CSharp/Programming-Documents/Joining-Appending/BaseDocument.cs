using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class BaseDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            //ExStart
            //ExId:AppendDocument_BaseDocument
            //ExSummary:Shows how to remove all content from a document before using it as a base to append documents to.
            // Use a blank document as the destination document.
            Document dstDoc = new Document();
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // The destination document is not actually empty which often causes a blank page to appear before the appended document
            // This is due to the base document having an empty section and the new document being started on the next page.
            // Remove all content from the destination document before appending.
            dstDoc.RemoveAllChildren();

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.BaseDocument Out.doc");

            Console.WriteLine("\nDocument appended successfully with all content removed from the destination document.\nFile saved at " + dataDir + "TestFile.BaseDocument Out.doc");
        }
    }
}
