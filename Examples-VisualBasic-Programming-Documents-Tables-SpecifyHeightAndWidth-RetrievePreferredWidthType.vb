' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

' Retrieve the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
table.AllowAutoFit = True

Dim firstCell As Cell = table.FirstRow.FirstCell
Dim type As PreferredWidthType = firstCell.CellFormat.PreferredWidth.Type
Dim value As Double = firstCell.CellFormat.PreferredWidth.Value

