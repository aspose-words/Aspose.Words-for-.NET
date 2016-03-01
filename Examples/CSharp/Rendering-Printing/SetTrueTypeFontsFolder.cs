using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using Aspose.Words.Saving;
namespace CSharp.Rendering_and_Printing
{
    class SetTrueTypeFontsFolder
    {
        public static void Run()
        {
            //ExStart:SetTrueTypeFontsFolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); ;

            Document doc = new Document(dataDir + "Rendering.doc");

            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead.
            FontSettings.SetFontsFolder(@"C:\MyFonts\", false);
            dataDir = dataDir + "Rendering.SetFontsFolder_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:SetTrueTypeFontsFolder
            Console.WriteLine("\nTrue type fonts folder setup successfully.\nFile saved at " + dataDir);
        }
    }
}
