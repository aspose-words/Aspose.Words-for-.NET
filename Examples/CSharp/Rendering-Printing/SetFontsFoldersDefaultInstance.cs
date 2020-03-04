using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class SetFontsFoldersDefaultInstance
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            // ExStart:SetFontsFoldersDefaultInstance
            FontSettings.DefaultInstance.SetFontsFolder("C:\\MyFonts\\", true);
            // ExEnd:SetFontsFoldersDefaultInstance           

            Document doc = new Document(dataDir + "Rendering.doc");
            dataDir = dataDir + "Rendering.SetFontsFolders_out.pdf";
            doc.Save(dataDir);
        }
    }
}
