
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
    class HyphenateWordsOfLanguages
    {
        public static void Run()
        {
            //ExStart:HyphenateWordsOfLanguages
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Load the documents which store the shapes we want to render.
            Document doc = new Document(dataDir + "TestFile RenderShape.doc");
            Hyphenation.RegisterDictionary("en-US", dataDir + @"hyph_en_US.dic");
            Hyphenation.RegisterDictionary("de-CH", dataDir + @"hyph_de_CH.dic");

            dataDir = dataDir + "HyphenateWordsOfLanguages_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:HyphenateWordsOfLanguages
            Console.WriteLine("\nWords of special languages hyphenate successfully.\nFile saved at " + dataDir);
            
        }
        
    }
}
