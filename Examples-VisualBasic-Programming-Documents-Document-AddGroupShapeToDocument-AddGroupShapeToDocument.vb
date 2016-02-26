' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

Dim doc As New Document()
doc.EnsureMinimum()
Dim gs As New GroupShape(doc)

Dim shape As New Shape(doc, Aspose.Words.Drawing.ShapeType.AccentBorderCallout1)
shape.Width = 100
shape.Height = 100
gs.AppendChild(shape)

Dim shape1 As New Shape(doc, Aspose.Words.Drawing.ShapeType.ActionButtonBeginning)
shape1.Left = 100
shape1.Width = 100
shape1.Height = 200
gs.AppendChild(shape1)

gs.Width = 200
gs.Height = 200

gs.CoordSize = New System.Drawing.Size(200, 200)

Dim builder As New DocumentBuilder(doc)
builder.InsertNode(gs)


dataDir = dataDir & Convert.ToString("groupshape-doc_out_.doc")

' Save the document to disk.
doc.Save(dataDir)
