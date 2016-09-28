
using System;
using System.IO;
using Aspose.Words.Replacing;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class FindAndReplace
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "ReplaceSimple.doc";

            // Open the document.
            Document doc = new Document(dataDir + fileName);

            // Check the text of the document
            Console.WriteLine("Original document text: " + doc.Range.Text);

            // Replace the text in the document.
            doc.Range.Replace("_CustomerName_", "James Bond", new FindReplaceOptions(FindReplaceDirection.Forward));

            // Check the replacement was made.
            Console.WriteLine("Document text after replace: " + doc.Range.Text);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the modified document.
            doc.Save(dataDir);

            Console.WriteLine("\nText found and replaced successfully.\nFile saved at " + dataDir);
        }
    }
}
