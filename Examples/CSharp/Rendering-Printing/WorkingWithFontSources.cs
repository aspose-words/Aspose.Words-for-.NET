using Aspose.Words.Fonts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class WorkingWithFontSources
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            GetListOfAvailableFonts(dataDir);

        }

        public static void GetListOfAvailableFonts(string dataDir)
        {
            // ExStart:GetListOfAvailableFonts
            // The path to the documents directory.
            Document doc = new Document(dataDir + "TestFile.docx");

            FontSettings fontSettings = new FontSettings();
            ArrayList fontSources = new ArrayList(fontSettings.GetFontsSources());
             
            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts. 
            FolderFontSource folderFontSource = new FolderFontSource(dataDir, true);

            // Add the custom folder which contains our fonts to the list of existing font sources.
            fontSources.Add(folderFontSource);

            // Convert the Arraylist of source back into a primitive array of FontSource objects.
            FontSourceBase[] updatedFontSources = (FontSourceBase[])fontSources.ToArray(typeof(FontSourceBase));

            foreach (PhysicalFontInfo fontInfo in updatedFontSources[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : "+ fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                Console.WriteLine("Version  : " + fontInfo.Version);
                Console.WriteLine("FilePath : " + fontInfo.FilePath);
            }
             
            // ExEnd:GetListOfAvailableFonts      
        }
    }
}
