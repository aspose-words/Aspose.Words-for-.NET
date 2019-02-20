using Aspose.Words.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontSettings
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            EnableDisableFontSubstitution(dataDir);
            SetFontFallbackSettings(dataDir);
        }

        public static void EnableDisableFontSubstitution(string dataDir)
        {
            // ExStart:EnableDisableFontSubstitution
            // The path to the documents directory.
            Document doc = new Document(dataDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;

            // Set font settings
            doc.FontSettings = fontSettings;
            dataDir = dataDir + "Rendering.DisableFontSubstitution_out.pdf";
            doc.Save(dataDir);
            // ExEnd:EnableDisableFontSubstitution      
            Console.WriteLine("\nDocument is rendered to PDF with disabled font substitution.\nFile saved at " + dataDir);
        }

        public static void SetFontFallbackSettings(string dataDir)
        {
            // ExStart:SetFontFallbackSettings
            // The path to the documents directory.
            Document doc = new Document(dataDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(dataDir + "Fallback.xml");

            // Set font settings
            doc.FontSettings = fontSettings;
            dataDir = dataDir + "Rendering.FontFallback_out.pdf";
            doc.Save(dataDir);
            // ExEnd:SetFontFallbackSettings      
            Console.WriteLine("\nDocument is rendered to PDF with font fallback.\nFile saved at " + dataDir);
        }
    }
}
