' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

' This updates all fields in the document.
doc.UpdateFields()
dataDir = dataDir & Convert.ToString("Rendering.UpdateFields_out_.pdf")
doc.Save(dataDir)
