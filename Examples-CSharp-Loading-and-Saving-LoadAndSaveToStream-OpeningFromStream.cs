// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_QuickStart();
string fileName = "Document.doc";

// Open the stream. Read only access is enough for Aspose.Words to load a document.
Stream stream = File.OpenRead(dataDir + fileName);

// Load the entire document into memory.
Document doc = new Document(stream);

// You can close the stream now, it is no longer needed because the document is in memory.
stream.Close();
