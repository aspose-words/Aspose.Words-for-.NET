// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Get the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Get the second column in the table.
Column column = Column.FromIndex(table, 0);
// Print the plain text of the column to the screen.
Console.WriteLine(column.ToTxt());
// Create a new column to the left of this column.
// This is the same as using the "Insert Column Before" command in Microsoft Word.
Column newColumn = column.InsertColumnBefore();

// Add some text to each of the column cells.
foreach (Cell cell in newColumn.Cells)
    cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
