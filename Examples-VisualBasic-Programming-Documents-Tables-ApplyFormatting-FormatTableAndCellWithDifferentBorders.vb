' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

Dim table As Table = builder.StartTable()
builder.InsertCell()

' Set the borders for the entire table.
table.SetBorders(LineStyle.[Single], 2.0, Color.Black)
' Set the cell shading for this cell.
builder.CellFormat.Shading.BackgroundPatternColor = Color.Red
builder.Writeln("Cell #1")

builder.InsertCell()
' Specify a different cell shading for the second cell.
builder.CellFormat.Shading.BackgroundPatternColor = Color.Green
builder.Writeln("Cell #2")

' End this row.
builder.EndRow()

' Clear the cell formatting from previous operations.
builder.CellFormat.ClearFormatting()

' Create the second row.
builder.InsertCell()

' Create larger borders for the first cell of this row. This will be different.
' compared to the borders set for the table.
builder.CellFormat.Borders.Left.LineWidth = 4.0
builder.CellFormat.Borders.Right.LineWidth = 4.0
builder.CellFormat.Borders.Top.LineWidth = 4.0
builder.CellFormat.Borders.Bottom.LineWidth = 4.0
builder.Writeln("Cell #3")

builder.InsertCell()
' Clear the cell formatting from the previous cell.
builder.CellFormat.ClearFormatting()
builder.Writeln("Cell #4")
' Save finished document.
doc.Save(dataDir & Convert.ToString("Table.SetBordersAndShading_out_.doc"))
