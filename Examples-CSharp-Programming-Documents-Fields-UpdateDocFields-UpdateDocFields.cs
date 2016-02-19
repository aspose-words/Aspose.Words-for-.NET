// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

Document doc = new Document(dataDir + "Rendering.doc");

// This updates all fields in the document.
doc.UpdateFields();
dataDir = dataDir + "Rendering.UpdateFields_out_.pdf";
doc.Save(dataDir);
