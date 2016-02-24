' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
Dim fileName As String = "TestFile.doc"
' Open the document
Dim doc As New Document(dataDir & fileName)

Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

' Autofit the first table to the page width.
table.AutoFit(AutoFitBehavior.AutoFitToWindow)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the document to disk.
doc.Save(dataDir)
'ExEnd

Debug.Assert(doc.FirstSection.Body.Tables(0).PreferredWidth.Type = PreferredWidthType.Percent, "PreferredWidth type is not percent")
Debug.Assert(doc.FirstSection.Body.Tables(0).PreferredWidth.Value = 100, "PreferredWidth value is different than 100")
