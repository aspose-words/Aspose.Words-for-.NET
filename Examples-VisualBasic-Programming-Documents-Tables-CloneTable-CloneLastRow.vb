' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

' Retrieve the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Clone the last row in the table.
Dim clonedRow As Row = DirectCast(table.LastRow.Clone(True), Row)

' Remove all content from the cloned row's cells. This makes the row ready for
' new content to be inserted into.
For Each cell As Cell In clonedRow.Cells
    cell.RemoveAllChildren()
Next

' Add the row to the end of the table.
table.AppendChild(clonedRow)

dataDir = dataDir & Convert.ToString("Table.AddCloneRowToTable_out_.doc")
' Save the document to disk.
doc.Save(dataDir)
