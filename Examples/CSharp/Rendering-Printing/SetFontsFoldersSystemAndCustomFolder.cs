
using System;
using Aspose.Words.Fonts;
using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontsFoldersSystemAndCustomFolder
    {
        public static void Run()
        {
            // ExStart:SetFontsFoldersSystemAndCustomFolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            SetFontsFolders(dataDir);
            GetSubstitutionWithoutSuffixes(dataDir);
        }

        public static void SetFontsFolders(string dataDir)
        {
            // ExStart:SetFontsFoldersSystemAndCustomFolder
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
            dataDir = dataDir + "Rendering.SetFontsFolders_out.pdf";
            doc.Save(dataDir);
            // ExEnd:SetFontsFoldersSystemAndCustomFolder 
            Console.WriteLine("\nFonts system and coustom folder is setup.\nFile saved at " + dataDir);
        }

        //ExStart:GetSubstitutionWithoutSuffixes
        public static void GetSubstitutionWithoutSuffixes(string dataDir)
        {
            Document doc = new Document(dataDir + "Get substitution without suffixes.docx");

            DocumentSubstitutionWarnings substitutionWarningHandler = new DocumentSubstitutionWarnings();
            doc.WarningCallback = substitutionWarningHandler;

            ArrayList fontSources = new ArrayList(FontSettings.DefaultInstance.GetFontsSources());
            FolderFontSource folderFontSource = new FolderFontSource(dataDir + "Fonts/", true);
            fontSources.Add(folderFontSource);
    
            FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.ToArray(typeof(FontSourceBase));
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);
    
            doc.Save(dataDir + "GetSubstitutionWithoutSuffixes_out.pdf");

            Assert.AreEqual(
                "Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
                substitutionWarningHandler.FontWarnings[0].Description);
        }

        public class DocumentSubstitutionWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                    FontWarnings.Warning(info);
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection();
        }
        //ExEnd:GetSubstitutionWithoutSuffixes
    }
}
