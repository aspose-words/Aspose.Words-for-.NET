// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

// Load the document from disk.
Document doc = new Document(dataDir + "TestFile.doc");

Document clone = doc.Clone();

dataDir = dataDir + "TestFile_clone_out_.doc";

// Save the document to disk.
clone.Save(dataDir);
