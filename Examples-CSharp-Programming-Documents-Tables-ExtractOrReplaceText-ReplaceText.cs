// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir);

// Get the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Replace any instances of our string in the entire table.
table.Range.Replace("Carrots", "Eggs", true, true);
// Replace any instances of our string in the last cell of the table only.
table.LastRow.LastCell.Range.Replace("50", "20", true, true);

dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.ReplaceCellText_out_.doc";
doc.Save(dataDir); 
