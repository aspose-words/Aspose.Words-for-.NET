' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

' Retrieve the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Create a clone of the table.
Dim tableClone As Table = DirectCast(table.Clone(True), Table)

' Insert the cloned table into the document after the original
table.ParentNode.InsertAfter(tableClone, table)

' Insert an empty paragraph between the two tables or else they will be combined into one
' upon save. This has to do with document validation.
table.ParentNode.InsertAfter(New Paragraph(doc), table)
dataDir = dataDir & Convert.ToString("Table.CloneTableAndInsert_out_.doc")

' Save the document to disk.
doc.Save(dataDir)
