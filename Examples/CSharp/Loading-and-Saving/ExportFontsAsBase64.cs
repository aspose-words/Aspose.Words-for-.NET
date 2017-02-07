using System;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ExportFontsAsBase64
    {
        public static void Run()
        {
            // ExStart:ExportFontsAsBase64            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            string fileName = "Document.doc";
            Document doc = new Document(dataDir + fileName);
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;           
            dataDir = dataDir + "ExportFontsAsBase64_out.html";
            doc.Save(dataDir, saveOptions);
            // ExEnd:ExportFontsAsBase64
            Console.WriteLine("\nSave option specified successfully.\nFile saved at " + dataDir);
        }
    }
}
