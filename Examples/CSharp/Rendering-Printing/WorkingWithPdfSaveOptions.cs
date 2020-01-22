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
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            EscapeUriInPdf(dataDir);
            ExportHeaderFooterBookmarks(dataDir);
            ScaleWmfFontsToMetafileSize(dataDir);
            AdditionalTextPositioning(dataDir);
            ConversionToPDF17(dataDir);
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

        public static void ScaleWmfFontsToMetafileSize(String dataDir)
        {
            // ExStart:ScaleWmfFontsToMetafileSize
            // The path to the documents directory.
            Document doc = new Document(dataDir + "MetafileRendering.docx");

            MetafileRenderingOptions metafileRenderingOptions =
                       new MetafileRenderingOptions
                       {
                           ScaleWmfFontsToMetafileSize = false
                       };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            PdfSaveOptions options = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            dataDir = dataDir + "ScaleWmfFontsToMetafileSize_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:ScaleWmfFontsToMetafileSize
            Console.WriteLine("\nFonts as metafile are rendered to its default size in PDF. File saved at " + dataDir);
        }
        
        public static void AdditionalTextPositioning(string dataDir)
        {
            // ExStart:AdditionalTextPositioning
            // The path to the documents directory.
            Document doc = new Document(dataDir + "TestFile.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.AdditionalTextPositioning = true;

            dataDir = dataDir + "AdditionalTextPositioning_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:AdditionalTextPositioning
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void ConversionToPDF17(string dataDir)
        {
            // ExStart:ConversionToPDF17
            // The path to the documents directory.
            Document originalDoc = new Document(dataDir + "Document.docx");

            // Provide PDFSaveOption compliance to PDF17
            // or just convert without SaveOptions
            PdfSaveOptions pso = new PdfSaveOptions();
            pso.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(dataDir + "Output.pdf", pso);
            // ExEnd:ConversionToPDF17
            Console.WriteLine("\nFile saved at " + dataDir);
        }
    }
}
