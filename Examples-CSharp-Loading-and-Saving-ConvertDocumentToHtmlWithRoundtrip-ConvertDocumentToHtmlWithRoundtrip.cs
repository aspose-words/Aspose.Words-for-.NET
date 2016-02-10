// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
