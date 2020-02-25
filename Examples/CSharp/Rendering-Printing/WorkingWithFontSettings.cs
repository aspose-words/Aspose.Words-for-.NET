using Aspose.Words.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class WorkingWithFontSettings
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            FontSettingsWithLoadOption(dataDir);
        }

        public static void FontSettingsWithLoadOption(string dataDir)
        {
            // ExStart:FontSettingsWithLoadOption
            FontSettings fontSettings = new FontSettings();
            // init font settings
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(dataDir + "MyDocument.docx", loadOptions);

            LoadOptions loadOptions2 = new LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(dataDir + "MyDocument.docx", loadOptions2);
            // ExEnd:FontSettingsWithLoadOption   
        }

        public static void FontSettingsDefaultInstance(string dataDir)
        {
            // ExStart:FontSettingsFontSource
            // ExStart:FontSettingsDefaultInstance
            FontSettings fontSettings = FontSettings.DefaultInstance;
            // ExEnd:FontSettingsDefaultInstance   
            fontSettings.SetFontsSources(new FontSourceBase[]
             {
                 new SystemFontSource(),
                 new FolderFontSource("/home/user/MyFonts", true)
             });
            // ExEnd:FontSettingsFontSource

            // init font settings
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc1 = new Document(dataDir + "MyDocument.docx", loadOptions);

            LoadOptions loadOptions2 = new LoadOptions();
            loadOptions2.FontSettings = fontSettings;
            Document doc2 = new Document(dataDir + "MyDocument.docx", loadOptions2);
        }
    }
}
