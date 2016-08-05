
using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Fonts;
using Aspose.Words;
using System.Collections;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontsFoldersSystemAndCustomFolder
    {
        public static void Run()
        {
            //ExStart:SetFontsFoldersSystemAndCustomFolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "Rendering.doc");
            FontSettings FontSettings = new FontSettings();

            // Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
            // We add this array to a new ArrayList to make adding or removing font entries much easier.
            ArrayList fontSources = new ArrayList(FontSettings.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
            FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

            // Add the custom folder which contains our fonts to the list of existing font sources.
            fontSources.Add(folderFontSource);

            // Convert the Arraylist of source back into a primitive array of FontSource objects.
            FontSourceBase[] updatedFontSources = (FontSourceBase[])fontSources.ToArray(typeof(FontSourceBase));
            
            // Apply the new set of font sources to use.
            FontSettings.SetFontsSources(updatedFontSources);
            // Set font settings
            doc.FontSettings = FontSettings;
            dataDir = dataDir + "Rendering.SetFontsFolders_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:SetFontsFoldersSystemAndCustomFolder 
            Console.WriteLine("\nFonts system and coustom folder is setup.\nFile saved at " + dataDir);
                     
        }
    }
}
