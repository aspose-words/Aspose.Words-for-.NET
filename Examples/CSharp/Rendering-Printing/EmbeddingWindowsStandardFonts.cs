using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class EmbeddingWindowsStandardFonts
    {
        public static void Run()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");

            // To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
            PdfSaveOptions options = new PdfSaveOptions();
            options.UseCoreFonts = true;

            string outPath = dataDir + "Rendering.DisableEmbedWindowsFonts_out_.pdf";
            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            doc.Save(outPath);
            //ExEnd:AvoidEmbeddingCoreFonts
            Console.WriteLine("\nAvoid embedded core fonts setup successfully.\nFile saved at " + outPath);
            SkipEmbeddedArialAndTimesRomanFonts(doc, dataDir);
        }
        private static void SkipEmbeddedArialAndTimesRomanFonts(Document doc, string dataDir)
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

            dataDir = dataDir + "Rendering.DisableEmbedWindowsFonts_out_.pdf";
            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(dataDir);
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
            Console.WriteLine("\nEmbedded arial and times new roman fonts are skipped successfully.\nFile saved at " + dataDir);
        }
    }
}
