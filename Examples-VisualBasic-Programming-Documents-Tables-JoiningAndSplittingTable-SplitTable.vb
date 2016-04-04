' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Load the document.
Dim doc As New Document(dataDir & fileName)

' Get the first table in the document.
Dim firstTable As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' We will split the table at the third row (inclusive).
Dim row As Row = firstTable.Rows(2)

' Create a new container for the split table.
Dim table As Table = DirectCast(firstTable.Clone(False), Table)

' Insert the container after the original.
firstTable.ParentNode.InsertAfter(table, firstTable)

' Add a buffer paragraph to ensure the tables stay apart.
firstTable.ParentNode.InsertAfter(New Paragraph(doc), firstTable)

Dim currentRow As Row

Do
    currentRow = firstTable.LastRow
    table.PrependChild(currentRow)
Loop While currentRow IsNot row

dataDir = dataDir & Convert.ToString("Table.SplitTable_out_.doc")
' Save the finished document.
doc.Save(dataDir)
