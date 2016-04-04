// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.EmptyTable.doc");

Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
// Clear any existing borders from the table.
table.ClearBorders();
// Set a green border around and inside the table.
table.SetBorders(LineStyle.Single, 1.5, Color.Green);

dataDir = dataDir + "Table.SetAllBorders_out_.doc";
// Save the document to disk.
doc.Save(dataDir);
