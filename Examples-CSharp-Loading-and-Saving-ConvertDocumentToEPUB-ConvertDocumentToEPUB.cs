// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

// Load the document from disk.
Document doc = new Document(dataDir + "Test File (doc).doc");

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

