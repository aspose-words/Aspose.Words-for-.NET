// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.SimpleTable.doc");

// Retrieve the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Clone the last row in the table.
Row clonedRow = (Row)table.LastRow.Clone(true);

// Remove all content from the cloned row's cells. This makes the row ready for
// new content to be inserted into.
foreach (Cell cell in clonedRow.Cells)
    cell.RemoveAllChildren();

// Add the row to the end of the table.
table.AppendChild(clonedRow);

dataDir = dataDir + "Table.AddCloneRowToTable_out_.doc";
// Save the document to disk.
doc.Save(dataDir);
