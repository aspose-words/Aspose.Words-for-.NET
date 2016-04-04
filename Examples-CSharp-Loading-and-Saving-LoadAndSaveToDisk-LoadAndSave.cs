// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_QuickStart();
string fileName = "Document.doc";
// Load the document from the absolute path on disk.
Document doc = new Document(dataDir + fileName);
dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the document as DOCX document.");
doc.Save(dataDir);
