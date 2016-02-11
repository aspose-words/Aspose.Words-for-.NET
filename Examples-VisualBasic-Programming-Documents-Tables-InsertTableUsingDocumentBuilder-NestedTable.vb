' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Build the outer table.
Dim cell As Cell = builder.InsertCell()
builder.Writeln("Outer Table Cell 1")

builder.InsertCell()
builder.Writeln("Outer Table Cell 2")

' This call is important in order to create a nested table within the first table
' Without this call the cells inserted below will be appended to the outer table.
builder.EndTable()

' Move to the first cell of the outer table.
builder.MoveTo(cell.FirstParagraph)

' Build the inner table.
builder.InsertCell()
builder.Writeln("Inner Table Cell 1")
builder.InsertCell()
builder.Writeln("Inner Table Cell 2")
builder.EndTable()

dataDir = dataDir & Convert.ToString("DocumentBuilder.InsertNestedTable_out_.doc")
' Save the document to disk.
doc.Save(dataDir)
