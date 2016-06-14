using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class EmbeddedFontsInPDF
    {
        public static void Run()
        {
            //ExStart:EmbeddAllFonts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // each time a document is rendered.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            string outPath = dataDir + "Rendering.EmbedFullFonts_out_.pdf";
            // The output PDF will be embedded with all fonts found in the document.
            doc.Save(outPath, options);
            //ExEnd:EmbeddAllFonts
            Console.WriteLine("\nAll Fonts embedded successfully.\nFile saved at " + outPath);
            EmbeddSubsetFonts(doc, dataDir);
        }
        private static void EmbeddSubsetFonts(Document doc, string dataDir)
        {
            //ExStart:EmbeddSubsetFonts
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;
            dataDir = dataDir + "Rendering.SubsetFonts_out_.pdf";
            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // in the document are included in the PDF fonts.
            doc.Save(dataDir, options);
            //ExEnd:EmbeddSubsetFonts
            Console.WriteLine("\nSubset Fonts embedded successfully.\nFile saved at " + dataDir);
        }
    }
}
