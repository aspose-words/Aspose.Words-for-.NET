' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document(dataDir & Convert.ToString("Table.TableStyle.docx"))

' Get the first cell of the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
Dim firstCell As Cell = table.FirstRow.FirstCell

' First print the color of the cell shading. This should be empty as the current shading
' is stored in the table style.
Dim cellShadingBefore As Color = firstCell.CellFormat.Shading.BackgroundPatternColor
Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore.ToString())

' Expand table style formatting to direct formatting.
doc.ExpandTableStylesToDirectFormatting()

' Now print the cell shading after expanding table styles. A blue background pattern color
' should have been applied from the table style.
Dim cellShadingAfter As Color = firstCell.CellFormat.Shading.BackgroundPatternColor
Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter.ToString())
