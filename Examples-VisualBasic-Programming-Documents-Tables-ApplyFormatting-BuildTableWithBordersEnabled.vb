' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.EmptyTable.doc"))

Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
' Clear any existing borders from the table.
table.ClearBorders()
' Set a green border around and inside the table.
table.SetBorders(LineStyle.[Single], 1.5, Color.Green)

dataDir = dataDir & Convert.ToString("Table.SetAllBorders_out_.doc")
' Save the document to disk.
doc.Save(dataDir)
