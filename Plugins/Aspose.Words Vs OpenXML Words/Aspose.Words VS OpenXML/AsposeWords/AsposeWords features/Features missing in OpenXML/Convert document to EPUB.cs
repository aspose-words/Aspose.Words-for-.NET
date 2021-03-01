// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ConvertDocumentToEpub : TestUtil
    {
        [Test]
        public void ConvertDocumentToEpubFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

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
            doc.Save(ArtifactsDir + "Convert document to EPUB - Aspose.Words.epub", saveOptions);
        }
    }
}
