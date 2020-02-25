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
            // Get available system fonts
            foreach (PhysicalFontInfo fontInfo in new SystemFontSource().GetAvailableFonts())
            {
                Console.WriteLine("\nFontFamilyName : " + fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                Console.WriteLine("Version  : " + fontInfo.Version);
                Console.WriteLine("FilePath : " + fontInfo.FilePath);
            }

            // Get available fonts in folder
            foreach (PhysicalFontInfo fontInfo in new FolderFontSource(dataDir, true).GetAvailableFonts())
            {
                Console.WriteLine("\nFontFamilyName : " + fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                Console.WriteLine("Version  : " + fontInfo.Version);
                Console.WriteLine("FilePath : " + fontInfo.FilePath);
            }

            // Get available fonts from FontSettings
            foreach (FontSourceBase fontsSource in FontSettings.DefaultInstance.GetFontsSources())
            {
                foreach (PhysicalFontInfo fontInfo in fontsSource.GetAvailableFonts())
                {
                    Console.WriteLine("\nFontFamilyName : " + fontInfo.FontFamilyName);
                    Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                    Console.WriteLine("Version  : " + fontInfo.Version);
                    Console.WriteLine("FilePath : " + fontInfo.FilePath);
                }
            }
            // ExEnd:GetListOfAvailableFonts      
        }
    }
}
