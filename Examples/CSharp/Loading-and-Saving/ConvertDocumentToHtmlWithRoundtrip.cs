
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToHtmlWithRoundtrip
    {
        public static void Run()
        {
            //ExStart:ConvertDocumentToHtmlWithRoundtrip
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (doc).doc");

            HtmlSaveOptions options = new HtmlSaveOptions();

            //HtmlSaveOptions.ExportRoundtripInformation property specifies
            //whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
            //Default value is true for HTML and false for MHTML and EPUB.
            options.ExportRoundtripInformation = true;
            
            doc.Save(dataDir + "ExportRoundtripInformation_out_.html", options);
            //ExEnd:ConvertDocumentToHtmlWithRoundtrip

            Console.WriteLine("\nDocument converted to html with roundtrip informations successfully.");
        }
    }
}
