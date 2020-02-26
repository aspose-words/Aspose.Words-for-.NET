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
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            
            EmbeddAllFonts(dataDir);
            EmbeddSubsetFonts(dataDir);
            SetFontEmbeddingMode(dataDir);
        }

        private static void EmbeddAllFonts(string dataDir)
        {
            // ExStart:EmbeddAllFonts
            // Load the document to render.
            Document doc = new Document(dataDir + "Rendering.doc");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // Each time a document is rendered.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            string outPath = dataDir + "Rendering.EmbedFullFonts_out.pdf";

            // The output PDF will be embedded with all fonts found in the document.
            doc.Save(outPath, options);
            // ExEnd:EmbeddAllFonts
            Console.WriteLine("\nAll Fonts embedded successfully.\nFile saved at " + outPath);
        }

        private static void EmbeddSubsetFonts(string dataDir)
        {
            // ExStart:EmbeddSubsetFonts
            // Load the document to render.
            Document doc = new Document(dataDir + "Rendering.doc");

            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;
            
            dataDir = dataDir + "Rendering.SubsetFonts_out.pdf";

            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // In the document are included in the PDF fonts.
            doc.Save(dataDir, options);
            // ExEnd:EmbeddSubsetFonts
            Console.WriteLine("\nSubset Fonts embedded successfully.\nFile saved at " + dataDir);
        }

        private static void SetFontEmbeddingMode(string dataDir)
        {
            // ExStart:SetFontEmbeddingMode
            // Load the document to render.
            Document doc = new Document(dataDir + "Rendering.doc");

            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone;

            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(dataDir + "Rendering.DisableEmbedWindowsFonts.pdf");
            // ExEnd:SetFontEmbeddingMode
            Console.WriteLine("\n Fonts embedding mode set successfully.\nFile saved at " + dataDir);
        }
    }
}
