// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.Document.doc");
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            
// Retrieve the first row in the table.
Row firstRow = table.FirstRow;
// Modify some row level properties.
firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
firstRow.RowFormat.HeightRule = HeightRule.Auto;
firstRow.RowFormat.AllowBreakAcrossPages = true; 
