
using System;
using System.IO;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing.Imaging;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class LoadHyphenationDictionaryForLanguage
    {
        public static void Run()
        {
            //ExStart:LoadHyphenationDictionaryForLanguage
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Load the documents which store the shapes we want to render.
            Document doc = new Document(dataDir + "TestFile RenderShape.doc");
            Stream stream = File.OpenRead(dataDir + @"hyph_de_CH.dic");
            Hyphenation.RegisterDictionary("de-CH", stream);

            dataDir = dataDir + "LoadHyphenationDictionaryForLanguage_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:LoadHyphenationDictionaryForLanguage
            Console.WriteLine("\nHyphenation dictionary for special language loaded successfully.\nFile saved at " + dataDir);           
        }
        
    }
}
