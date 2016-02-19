// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Load the document.
Document doc = new Document(dataDir + fileName);

// Get the first and second table in the document.
// The rows from the second table will be appended to the end of the first table.
Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

// Append all rows from the current table to the next.
// Due to the design of tables even tables with different cell count and widths can be joined into one table.
while (secondTable.HasChildNodes)
    firstTable.Rows.Add(secondTable.FirstRow);

// Remove the empty table container.
secondTable.Remove();
dataDir = dataDir + "Table.CombineTables_out_.doc";
// Save the finished document.
doc.Save(dataDir);
