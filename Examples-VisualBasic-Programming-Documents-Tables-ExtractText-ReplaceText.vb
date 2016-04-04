' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir)

' Get the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Replace any instances of our string in the entire table.
table.Range.Replace("Carrots", "Eggs", True, True)
' Replace any instances of our string in the last cell of the table only.
table.LastRow.LastCell.Range.Replace("50", "20", True, True)

dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.ReplaceCellText_out_.doc"
doc.Save(dataDir)
