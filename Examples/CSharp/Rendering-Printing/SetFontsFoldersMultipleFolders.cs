
using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontsFoldersMultipleFolders
    {
        public static void Run()
        {
            //ExStart:SetFontsFoldersMultipleFolders
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");
            FontSettings FontSettings = new FontSettings();

            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead.
            FontSettings.SetFontsFolders(new string[] { @"C:\MyFonts\", @"D:\Misc\Fonts\" }, true);
            // Set font settings
            doc.FontSettings = FontSettings;
            dataDir = dataDir + "Rendering.SetFontsFolders_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:SetFontsFoldersMultipleFolders           
        }
    }
}
