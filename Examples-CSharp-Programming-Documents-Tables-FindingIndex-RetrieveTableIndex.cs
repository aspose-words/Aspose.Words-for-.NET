// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Get the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
int tableIndex = allTables.IndexOf(table);            
