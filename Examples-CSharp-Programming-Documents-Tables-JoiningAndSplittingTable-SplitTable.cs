// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Load the document.
Document doc = new Document(dataDir + fileName);

// Get the first table in the document.
Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);

// We will split the table at the third row (inclusive).
Row row = firstTable.Rows[2];

// Create a new container for the split table.
Table table = (Table)firstTable.Clone(false);

// Insert the container after the original.
firstTable.ParentNode.InsertAfter(table, firstTable);

// Add a buffer paragraph to ensure the tables stay apart.
firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

Row currentRow;

do
{
    currentRow = firstTable.LastRow;
    table.PrependChild(currentRow);
}
while (currentRow != row);

dataDir = dataDir + "Table.SplitTable_out_.doc";
// Save the finished document.
doc.Save(dataDir);
