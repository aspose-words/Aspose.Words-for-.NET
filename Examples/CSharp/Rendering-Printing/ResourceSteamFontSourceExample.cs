using Aspose.Words.Fonts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    // ExStart:ResourceSteamFontSourceExample
    class ResourceSteamFontSourceExample : StreamFontSource
    {
        public override Stream OpenFontDataStream()
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream("resourceName");
        }

        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            Document doc = new Document(dataDir + "Rendering.doc");

            // FontSettings.SetFontSources instead.
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[] { new SystemFontSource(), new ResourceSteamFontSourceExample() });
            doc.Save(dataDir + "Rendering.SetFontsFolders.pdf");
        }
    }
    // ExStart:ResourceSteamFontSourceExample
}
