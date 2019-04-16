using Aspose.Words.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class WorkingWithFontResolution
    {
        public static void Run()
        {
            FontSettingsWithLoadOptions();
            SetFontsFolder();
        }

        static void FontSettingsWithLoadOptions()
        {
            // ExStart:FontSettingsWithLoadOptions
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            FontSettings fontSettings = new FontSettings();
            TableSubstitutionRule substitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS".
            substitutionRule.AddSubstitutes("UnknownFont1", new string[] { "Comic Sans MS" });
            LoadOptions lo = new LoadOptions();
            lo.FontSettings = fontSettings;
            Document doc = new Document(dataDir + "myfile.html", lo);
            // ExEnd:FontSettingsWithLoadOptions
            Console.WriteLine("\nFile created successfully.\nFile saved at " + dataDir);
        }

        static void SetFontsFolder()
        {
            // ExStart:SetFontsFolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(dataDir + "Fonts", false);
            LoadOptions lo = new LoadOptions();
            lo.FontSettings = fontSettings;
            Document doc = new Document(dataDir + "myfile.html", lo);
            // ExEnd:SetFontsFolder
        }
    }
}
