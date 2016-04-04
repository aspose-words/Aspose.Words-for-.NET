// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.SimpleTable.doc");

// Retrieve the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
table.AllowAutoFit = true;

Cell firstCell = table.FirstRow.FirstCell;
PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
double value = firstCell.CellFormat.PreferredWidth.Value;

