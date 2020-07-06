using System;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ExportResourcesUsingHtmlSaveOptions
    {
        public static void Run()
        {
            // ExStart:ExportResourcesUsingHtmlSaveOptions            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            string fileName = "Document.docx";
            Document doc = new Document(dataDir + fileName);
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = dataDir + @"\Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";
            doc.Save(dataDir + "ExportResourcesUsingHtmlSaveOptions.html", saveOptions);
            // ExEnd:ExportResourcesUsingHtmlSaveOptions
            Console.WriteLine("\nSave option specified successfully.\nFile saved at " + dataDir);
        }
    }
}
