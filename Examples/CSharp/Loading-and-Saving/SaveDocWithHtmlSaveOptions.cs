using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveDocWithHtmlSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            SaveHtmlWithMetafileFormat(dataDir); 
        }

        public static void SaveHtmlWithMetafileFormat(string dataDir)
        {
            // ExStart:SaveHtmlWithMetafileFormat
            Document doc = new Document(dataDir + "Document.docx");
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.EmfOrWmf;

            dataDir = dataDir + "SaveHtmlWithMetafileFormat_out.html";
            doc.Save(dataDir, options);
            // ExEnd:SaveHtmlWithMetafileFormat
            Console.WriteLine("\nDocument saved with Metafile format.\nFile saved at " + dataDir);
        }
    }
}
