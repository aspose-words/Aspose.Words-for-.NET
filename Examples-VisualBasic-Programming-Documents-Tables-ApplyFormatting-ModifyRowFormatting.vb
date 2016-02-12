' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.Document.doc"))
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Retrieve the first row in the table.
Dim firstRow As Row = table.FirstRow
' Modify some row level properties.
firstRow.RowFormat.Borders.LineStyle = LineStyle.None
firstRow.RowFormat.HeightRule = HeightRule.Auto
firstRow.RowFormat.AllowBreakAcrossPages = True
