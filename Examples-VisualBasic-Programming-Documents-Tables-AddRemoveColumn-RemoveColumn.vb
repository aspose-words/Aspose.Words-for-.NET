' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Get the second table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 1, True), Table)

' Get the third column from the table and remove it.
Dim column__1 As Column = Column.FromIndex(table, 2)
column__1.Remove()
