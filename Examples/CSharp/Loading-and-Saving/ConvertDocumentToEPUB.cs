
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Saving;
namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToEPUB
    {
        public static void Run()
        {
            //ExStart:ConvertDocumentToEPUB
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Document.EpubConversion.doc");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved.
            HtmlSaveOptions saveOptions =
                new HtmlSaveOptions();

            // Specify the desired encoding.
            saveOptions.Encoding = System.Text.Encoding.UTF8;

            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Specify that we want to export document properties.
            saveOptions.ExportDocumentProperties = true;

            // Specify that we want to save in EPUB format.
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file.
            doc.Save(dataDir + "Document.EpubConversion_out_.epub", saveOptions);
            //ExEnd:ConvertDocumentToEPUB
            ConvertDocumentToEPUBUsingDefaultSaveOption();
            Console.WriteLine("\nDocument converted to EPUB successfully.");
        }
        //ExStart:ConvertDocumentToEPUBUsingDefaultSaveOption
        public static void ConvertDocumentToEPUBUsingDefaultSaveOption()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Document.EpubConversion.doc");
            // Save the document in EPUB format.
            doc.Save(dataDir + "Document.EpubConversionUsingDefaultSaveOption_out_.epub");
        }
        //ExEnd:ConvertDocumentToEPUBUsingDefaultSaveOption
            
    }
}
