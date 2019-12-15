
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

        // ExStart:PdfRenderWarnings
        public void PdfRenderWarnings()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document from disk.
            Document doc = new Document(dataDir + "PdfRenderWarnings.doc");

            // Set a SaveOptions object to not emulate raster operations.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRasterOperations = false,
                RenderingMode = MetafileRenderingMode.VectorWithFallback
            };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(dataDir + "PdfRenderWarnings.pdf", saveOptions);

            // While the file saves successfully, rendering warnings that occurred during saving are collected here.
            foreach (WarningInfo warningInfo in callback.mWarnings)
            {
                Console.WriteLine(warningInfo.Description);
            }
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }
        // ExEnd:PdfRenderWarnings
    }
}
