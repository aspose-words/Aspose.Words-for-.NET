' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

Dim table As Table = builder.StartTable()
' We must insert at least one row first before setting any table formatting.
builder.InsertCell()
' Set the table style used based of the unique style identifier.
' Note that not all table styles are available when saving as .doc format.
table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1
' Apply which features should be formatted by the style.
table.StyleOptions = TableStyleOptions.FirstColumn Or TableStyleOptions.RowBands Or TableStyleOptions.FirstRow
table.AutoFit(AutoFitBehavior.AutoFitToContents)

' Continue with building the table as normal.
builder.Writeln("Item")
builder.CellFormat.RightPadding = 40
builder.InsertCell()
builder.Writeln("Quantity (kg)")
builder.EndRow()

builder.InsertCell()
builder.Writeln("Apples")
builder.InsertCell()
builder.Writeln("20")
builder.EndRow()

builder.InsertCell()
builder.Writeln("Bananas")
builder.InsertCell()
builder.Writeln("40")
builder.EndRow()

builder.InsertCell()
builder.Writeln("Carrots")
builder.InsertCell()
builder.Writeln("50")
builder.EndRow()

dataDir = dataDir & Convert.ToString("DocumentBuilder.SetTableStyle_out_.docx")

' Save the document to disk.
doc.Save(dataDir)
