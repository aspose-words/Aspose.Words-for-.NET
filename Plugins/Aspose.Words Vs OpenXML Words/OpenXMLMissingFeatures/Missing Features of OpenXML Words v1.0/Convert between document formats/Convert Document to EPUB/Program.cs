// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Saving;

namespace ConvertDocumentToEPUB
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            // Open an existing document from disk.
            Document doc = new Document(MyDir + "Converting Document.docx");

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
            doc.Save(MyDir + "Document.EpubConversion Out.epub", saveOptions);
        }

    }
}
