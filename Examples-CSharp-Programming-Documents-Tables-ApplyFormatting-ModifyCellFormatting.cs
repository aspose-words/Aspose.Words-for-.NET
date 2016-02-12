// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.Document.doc"); 
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Retrieve the first cell in the table.
Cell firstCell = table.FirstRow.FirstCell;
// Modify some cell level properties.
firstCell.CellFormat.Width = 30; // in points
firstCell.CellFormat.Orientation = TextOrientation.Downward;
firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;
