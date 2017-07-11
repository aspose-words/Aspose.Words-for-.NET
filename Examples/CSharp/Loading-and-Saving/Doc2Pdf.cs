
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class Doc2Pdf
    {
        public static void Run()
        {
            // ExStart:Doc2Pdf
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Template.doc");

            dataDir = dataDir + "Template_out.pdf";

            // Save the document in PDF format.
            doc.Save(dataDir);
            // ExEnd:Doc2Pdf
            Console.WriteLine("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
        }

        public static void DisplayDocTitleInWindowTitlebar()
        {
            // ExStart:DisplayDocTitleInWindowTitlebar
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Template.doc");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DisplayDocTitle = true;

            dataDir = dataDir + "Template_out.pdf";

            // Save the document in PDF format.
            doc.Save(dataDir, saveOptions);
            // ExEnd:DisplayDocTitleInWindowTitlebar
            Console.WriteLine("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
        }
    }
}
