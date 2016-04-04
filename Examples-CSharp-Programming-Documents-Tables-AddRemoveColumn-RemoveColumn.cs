// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Get the second table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 1, true);

// Get the third column from the table and remove it.
Column column = Column.FromIndex(table, 2);
column.Remove();
