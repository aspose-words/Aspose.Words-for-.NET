// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithTables();
string fileName = "TestFile.doc";
// Open the document
Document doc = new Document(dataDir + fileName);

Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Autofit the first table to the page width.
table.AutoFit(AutoFitBehavior.AutoFitToWindow);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName); 
// Save the document to disk.
doc.Save(dataDir);

Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Percent, "PreferredWidth type is not percent");
Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 100, "PreferredWidth value is different than 100");
