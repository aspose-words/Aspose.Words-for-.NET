using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class Doc2Pdf
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            SaveDoc2Pdf(dataDir);
            DisplayDocTitleInWindowTitlebar(dataDir);
            PdfRenderWarnings(dataDir);
        }

        public static void SaveDoc2Pdf(string dataDir)
        {
            // ExStart:Doc2Pdf
            // Load the document from disk.
            Document doc = new Document(dataDir + "Rendering.doc");

            // Save the document in PDF format.
            doc.Save(dataDir + "SaveDoc2Pdf.pdf");
            // ExEnd:Doc2Pdf

            Console.WriteLine("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
        }

        public static void DisplayDocTitleInWindowTitlebar(string dataDir)
        {
            // ExStart:DisplayDocTitleInWindowTitlebar
            // Load the document from disk.
            Document doc = new Document(dataDir + "Rendering.doc");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DisplayDocTitle = true;

            // Save the document in PDF format.
            doc.Save(dataDir + "DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
            // ExEnd:DisplayDocTitleInWindowTitlebar
            
            Console.WriteLine("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
        }

        // ExStart:PdfRenderWarnings
        public static void PdfRenderWarnings(string dataDir)
        {
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
