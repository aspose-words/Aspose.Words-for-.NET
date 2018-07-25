using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class WorkingWithPdfSaveOptions
    {
        public static void Run()
        {
            // ExStart:SetTrueTypeFontsFolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            EscapeUriInPdf(dataDir);
            ExportHeaderFooterBookmarks(dataDir);
        }

        public static void EscapeUriInPdf(String dataDir)
        {
            // ExStart:EscapeUriInPdf
            // The path to the documents directory.
            Document doc = new Document(dataDir + "EscapeUri.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = false;

            dataDir = dataDir + "EscapeUri_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:EscapeUriInPdf
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void ExportHeaderFooterBookmarks(String dataDir)
        {
            // ExStart:ExportHeaderFooterBookmarks
            // The path to the documents directory.
            Document doc = new Document(dataDir + "TestFile.docx");


            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            options.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            dataDir = dataDir + "ExportHeaderFooterBookmarks_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:ExportHeaderFooterBookmarks
            Console.WriteLine("\nFile saved at " + dataDir);
        }
    }
}
