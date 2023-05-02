// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
int tableIndex = allTables.IndexOf(table);
