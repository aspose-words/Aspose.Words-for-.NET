using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToHtml
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            ExportRoundtripInformation(dataDir);
            SplitDocumentByHeadingsHtml(dataDir);
            SplitDocumentBySectionsHtml(dataDir);
        }

        public static void ExportRoundtripInformation(string dataDir)
        {
            // ExStart:ConvertDocumentToHtmlWithRoundtrip
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (doc).docx");

            HtmlSaveOptions options = new HtmlSaveOptions();

            // HtmlSaveOptions.ExportRoundtripInformation property specifies
            // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
            // Default value is true for HTML and false for MHTML and EPUB.
            options.ExportRoundtripInformation = true;
            
            doc.Save(dataDir + "ExportRoundtripInformation_out.html", options);
            // ExEnd:ConvertDocumentToHtmlWithRoundtrip

            Console.WriteLine("\nDocument converted to html with roundtrip informations successfully.");
        }

        public static void SplitDocumentByHeadingsHtml(string dataDir)
        {
            //ExStart:SplitDocumentByHeadingsHtml
            // Open a Word document
            Document doc = new Document(dataDir + "Test File (doc).docx");
 
            HtmlSaveOptions options = new HtmlSaveOptions();
            // Split a document into smaller parts, in this instance split by heading
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
 
            // Save the output file
            doc.Save(dataDir + "SplitDocumentByHeadings_out.html", options);
            //ExEnd:SplitDocumentByHeadingsHtml
        }

        public static void SplitDocumentBySectionsHtml(string dataDir)
        {
            // Open a Word document
            Document doc = new Document(dataDir + "Test File (doc).docx");
 
            //ExStart:SplitDocumentBySectionsHtml
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            //ExEnd:SplitDocumentBySectionsHtml
            
            // Save the output file
            doc.Save(dataDir + "SplitDocumentBySections_out.html", options);
            
        }
    }
}
