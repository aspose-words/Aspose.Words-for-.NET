' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Get the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

Dim allTables As NodeCollection = doc.GetChildNodes(NodeType.Table, True)
Dim tableIndex As Integer = allTables.IndexOf(table)
