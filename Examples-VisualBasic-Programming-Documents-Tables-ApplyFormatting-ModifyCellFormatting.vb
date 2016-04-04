' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.Document.doc"))
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Retrieve the first cell in the table.
Dim firstCell As Cell = table.FirstRow.FirstCell
' Modify some cell level properties.
firstCell.CellFormat.Width = 30
' in points
firstCell.CellFormat.Orientation = TextOrientation.Downward
firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen
