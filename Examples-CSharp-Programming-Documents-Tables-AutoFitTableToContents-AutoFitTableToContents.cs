// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithTables();

string fileName = "TestFile.doc";
Document doc = new Document(dataDir + fileName);

Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Auto fit the table to the cell contents
table.AutoFit(AutoFitBehavior.AutoFitToContents);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the document to disk.
doc.Save(dataDir);

Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto, "PreferredWidth type is not auto");
Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Type == PreferredWidthType.Auto, "PrefferedWidth on cell is not auto");
Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Value == 0, "PreferredWidth value is not 0");
