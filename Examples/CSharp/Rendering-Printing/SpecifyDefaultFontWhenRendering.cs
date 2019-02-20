
using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SpecifyDefaultFontWhenRendering
    {
        public static void Run()
        {
            // ExStart:SpecifyDefaultFontWhenRendering
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");

            FontSettings FontSettings = new FontSettings();

            // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
            FontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
            // Set font settings
            doc.FontSettings = FontSettings;
            dataDir = dataDir + "Rendering.SetDefaultFont_out.pdf";
            // Now the set default font is used in place of any missing fonts during any rendering calls.
            doc.Save(dataDir);            
            // ExEnd:SpecifyDefaultFontWhenRendering 
            Console.WriteLine("\nDefault font is setup during rendering.\nFile saved at " + dataDir);
        }
    }
}
